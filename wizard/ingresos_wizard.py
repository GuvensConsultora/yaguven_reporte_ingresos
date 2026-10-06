# -*- coding: utf-8 -*-
import io
from collections import defaultdict
from datetime import datetime, time, timedelta

from zoneinfo import ZoneInfo

from dateutil.relativedelta import relativedelta

from odoo import _, fields, models
from odoo.exceptions import UserError

from ..models.account_journal import TIPOS_INGRESO

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

# Orden y etiqueta de cada medio. Retenciones va aparte: no es plata que entra.
RETENCION = 'retention'
MEDIOS = [t for t in TIPOS_INGRESO if t[0] != RETENCION]
SIN_SUCURSAL = 'Sin sucursal'
# Estados de account.payment en Odoo 20 que NO son plata movida.
ESTADOS_FUERA = ('draft', 'canceled', 'rejected')
TZ_DEFECTO = 'America/Argentina/Buenos_Aires'


# Desde el menú «Reporte diario de ingresos» (contexto `yaguven_hoy`) las dos fechas son hoy;
# desde «Ingresos por período», el mes anterior completo.
def _default_desde(self):
    hoy = fields.Date.context_today(self)
    if self.env.context.get('yaguven_hoy'):
        return hoy
    return hoy.replace(day=1) - relativedelta(months=1)


def _default_hasta(self):
    hoy = fields.Date.context_today(self)
    if self.env.context.get('yaguven_hoy'):
        return hoy
    return hoy.replace(day=1) - relativedelta(days=1)


class IngresosWizard(models.TransientModel):
    _name = 'yaguven.ingresos.wizard'
    _description = 'Reporte de ingresos'

    # El mismo asistente sirve para el día y para el período (ver _default_desde).
    date_from = fields.Date(string='Desde', required=True, default=_default_desde)
    date_to = fields.Date(string='Hasta', required=True, default=_default_hasta)
    company_id = fields.Many2one(
        'res.company', string='Empresa', required=True,
        default=lambda self: self.env.company,
    )
    operating_unit_ids = fields.Many2many(
        'operating.unit', string='Sucursales',
        help='Vacío = todas las sucursales.',
    )

    # -------------------------------------------------------------------------
    # Motor
    # Qué entra:
    #  · pagos contables de CLIENTES (recibos, transferencias, cheques…). Los pagos a
    #    proveedores quedan afuera por construcción; los salientes a clientes son las
    #    reversiones (NC devueltas) y restan.
    #  · pagos de POS (lo que se cobra en el mostrador). Los pagos contables que el cierre
    #    de caja crea a partir de ellos (pos_session_id) se excluyen: si no, se cuentan dos
    #    veces.
    # -------------------------------------------------------------------------

    def _limites_utc(self):
        """Desde 00:00 hasta 24:00 hora local, en UTC (payment_date de POS es datetime)."""
        tz, utc = ZoneInfo(self.env.user.tz or TZ_DEFECTO), ZoneInfo('UTC')
        desde = datetime.combine(self.date_from, time.min, tz).astimezone(utc)
        hasta = datetime.combine(self.date_to + timedelta(days=1), time.min, tz).astimezone(utc)
        return desde.replace(tzinfo=None), hasta.replace(tzinfo=None)

    def _fila(self, fecha, recibo, pago, cliente, ou, diario, entra):
        tipo = diario.yaguven_ingreso_tipo or 'other'
        return {
            'fecha': fecha, 'recibo': recibo, 'pago': pago, 'cliente': cliente,
            'ou': ou, 'sucursal': ou.name if ou else SIN_SUCURSAL,
            'clave': tipo, 'diario': diario.name,
            'tipo': 'Cobro' if entra else 'Reversión',
        }

    def _filas_pagos(self):
        pagos = self.env['account.payment'].search([
            ('company_id', '=', self.company_id.id),
            ('partner_type', '=', 'customer'),
            ('state', 'not in', ESTADOS_FUERA),
            ('pos_session_id', '=', False),
            ('date', '>=', self.date_from),
            ('date', '<=', self.date_to),
        ], order='date, id')
        filas = []
        for p in pagos:
            entra = p.payment_type == 'inbound'
            f = self._fila(p.date, p.payment_group_id.display_name or '', p.name,
                           p.partner_id.display_name or '',
                           p.journal_id.operating_unit_id or p.move_id.operating_unit_id,
                           p.journal_id, entra)
            f['importe'] = p.amount if entra else -p.amount
            filas.append(f)
        return filas

    def _filas_pos(self):
        desde, hasta = self._limites_utc()
        pagos = self.env['pos.payment'].search([
            ('company_id', '=', self.company_id.id),
            ('payment_date', '>=', desde),
            ('payment_date', '<', hasta),
            # Sin diario = «Cuenta corriente»: no entra plata, queda como deuda del cliente.
            ('payment_method_id.journal_id', '!=', False),
        ], order='payment_date, id')
        filas = []
        for p in pagos:
            # El vuelto (is_change) es parte del cobro, no una devolución.
            entra = p.amount >= 0 or p.is_change
            f = self._fila(fields.Datetime.context_timestamp(self, p.payment_date).date(),
                           p.session_id.name or '', p.pos_order_id.name or '',
                           p.pos_order_id.partner_id.display_name or '',
                           p.session_id.config_id.operating_unit_id
                           or p.payment_method_id.journal_id.operating_unit_id,
                           p.payment_method_id.journal_id, entra)
            f['importe'] = p.amount
            filas.append(f)
        return filas

    def _datos(self):
        self.ensure_one()
        if self.date_from > self.date_to:
            raise UserError(_('La fecha «Desde» es posterior a «Hasta».'))
        # Todo en la empresa elegida, no en la activa del usuario: con otra activa, las
        # reglas multiempresa ocultaban los pagos y el reporte daba cero (O17, 05/10/2026).
        self = self.with_company(self.company_id)
        filas = self._filas_pagos() + self._filas_pos()
        if self.operating_unit_ids:
            filas = [f for f in filas if f['ou'] in self.operating_unit_ids]
        filas.sort(key=lambda f: (f['fecha'], f['sucursal']))

        etiquetas = dict(TIPOS_INGRESO)
        por_medio = defaultdict(lambda: {'cobro': 0.0, 'reversion': 0.0, 'n': 0})
        pivot = defaultdict(lambda: defaultdict(float))
        for f in filas:
            d = por_medio[f['clave']]
            d['cobro' if f['tipo'] == 'Cobro' else 'reversion'] += f['importe']
            d['n'] += 1
            if f['clave'] != RETENCION:
                pivot[f['sucursal']][f['clave']] += f['importe']
            f['medio'] = 'Retenciones' if f['clave'] == RETENCION else etiquetas[f['clave']]

        medios = [{'clave': c, 'medio': e, 'cobro': por_medio[c]['cobro'],
                   'reversion': por_medio[c]['reversion'],
                   'neto': por_medio[c]['cobro'] + por_medio[c]['reversion'],
                   'n': por_medio[c]['n']}
                  for c, e in MEDIOS if c in por_medio]
        claves = [m['clave'] for m in medios]
        sucursales = []
        for nombre in sorted(pivot, key=lambda s: (s == SIN_SUCURSAL, s)):
            vals = [pivot[nombre].get(c, 0.0) for c in claves]
            sucursales.append({'sucursal': nombre, 'valores': vals, 'total': sum(vals)})

        ret = por_medio.get(RETENCION, {'cobro': 0.0, 'n': 0})
        ncs = self.env['account.move'].search([
            ('company_id', '=', self.company_id.id), ('move_type', '=', 'out_refund'),
            ('state', '=', 'posted'),
            ('invoice_date', '>=', self.date_from), ('invoice_date', '<=', self.date_to)])
        prov = self.env['account.payment'].search([
            ('company_id', '=', self.company_id.id), ('partner_type', '=', 'supplier'),
            ('state', 'not in', ESTADOS_FUERA),
            ('date', '>=', self.date_from), ('date', '<=', self.date_to)])
        return {
            'empresa': self.company_id.name,
            'desde': self.date_from.strftime('%d/%m/%Y'),
            'hasta': self.date_to.strftime('%d/%m/%Y'),
            'filtro_sucursales': ', '.join(self.operating_unit_ids.mapped('name')),
            'medios': medios,
            'total_cobro': sum(m['cobro'] for m in medios),
            'total_reversion': sum(m['reversion'] for m in medios),
            'total_neto': sum(m['neto'] for m in medios),
            'columnas': [m['medio'] for m in medios],
            'sucursales': sucursales,
            'totales_sucursal': [sum(s['valores'][i] for s in sucursales)
                                 for i in range(len(claves))],
            'retenciones': ret['cobro'], 'retenciones_n': ret['n'],
            'nc': sum(ncs.mapped('amount_total_signed')), 'nc_n': len(ncs),
            'proveedores': sum(prov.mapped('amount')), 'proveedores_n': len(prov),
            'detalle': filas,
        }

    # -------------------------------------------------------------------------
    # Acciones
    # -------------------------------------------------------------------------

    def action_print_pdf(self):
        self.ensure_one()
        return self.env.ref('yaguven_reporte_ingresos.action_report_ingresos').report_action(self)

    def action_export_excel(self):
        self.ensure_one()
        if not xlsxwriter:
            raise UserError(_('Se requiere la librería xlsxwriter para exportar a Excel.'))
        datos = self._datos()
        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        self._write_excel(wb, datos)
        wb.close()
        nombre = 'ingresos_%s_%s.xlsx' % (self.date_from.strftime('%Y%m%d'),
                                          self.date_to.strftime('%Y%m%d'))
        adjunto = self.env['ir.attachment'].create({
            'name': nombre,
            'type': 'binary',
            # Odoo 20: el contenido va en `raw` (ya no existe `datas`).
            'raw': output.getvalue(),
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/%d?download=true' % adjunto.id,
            'target': 'new',
        }

    # -------------------------------------------------------------------------
    # Excel: Resumen, Por sucursal y Detalle (con filtro, para procesarlo a gusto)
    # -------------------------------------------------------------------------

    def _write_excel(self, wb, d):
        titulo = wb.add_format({'bold': True, 'font_size': 14})
        head = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#1F4E79',
                              'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        num = wb.add_format({'num_format': '#,##0.00;-#,##0.00'})
        tot_txt = wb.add_format({'bold': True, 'bg_color': '#DDEBF7'})
        tot_num = wb.add_format({'bold': True, 'bg_color': '#DDEBF7',
                                 'num_format': '#,##0.00;-#,##0.00'})
        fecha = wb.add_format({'num_format': 'dd/mm/yyyy'})

        ws = wb.add_worksheet('Resumen')
        ws.write(0, 0, 'Lo que entró del %s al %s — %s' % (d['desde'], d['hasta'], d['empresa']),
                 titulo)
        if d['filtro_sucursales']:
            ws.write(1, 0, 'Sucursales: %s' % d['filtro_sucursales'])
        for c, (txt, ancho) in enumerate([('Medio', 22), ('Cobrado', 18), ('Reversiones (NC)', 18),
                                          ('Neto que entró', 18), ('Cantidad', 10)]):
            ws.write(3, c, txt, head)
            ws.set_column(c, c, ancho)
        r = 4
        for m in d['medios']:
            ws.write(r, 0, m['medio'])
            ws.write_number(r, 1, m['cobro'], num)
            ws.write_number(r, 2, m['reversion'], num)
            ws.write_number(r, 3, m['neto'], num)
            ws.write_number(r, 4, m['n'])
            r += 1
        ws.write(r, 0, 'TOTAL', tot_txt)
        ws.write_number(r, 1, d['total_cobro'], tot_num)
        ws.write_number(r, 2, d['total_reversion'], tot_num)
        ws.write_number(r, 3, d['total_neto'], tot_num)
        ws.write(r, 4, '', tot_txt)
        r += 2
        for txt, imp, n in [
            ('Retenciones que nos hicieron (no es plata)', d['retenciones'], d['retenciones_n']),
            ('Notas de crédito a clientes del período', d['nc'], d['nc_n']),
            ('Pagos a proveedores (NO incluidos)', d['proveedores'], d['proveedores_n']),
        ]:
            ws.write(r, 0, txt)
            ws.write_number(r, 1, imp, num)
            ws.write_number(r, 4, n)
            r += 1

        ws2 = wb.add_worksheet('Por sucursal')
        cols = ['Sucursal'] + d['columnas'] + ['Total']
        for c, txt in enumerate(cols):
            ws2.write(0, c, txt, head)
            ws2.set_column(c, c, 22 if c == 0 else 16)
        r = 1
        for s in d['sucursales']:
            ws2.write(r, 0, s['sucursal'])
            for c, v in enumerate(s['valores'] + [s['total']], 1):
                ws2.write_number(r, c, v, num)
            r += 1
        ws2.write(r, 0, 'TOTAL', tot_txt)
        for c, v in enumerate(d['totales_sucursal'] + [d['total_neto']], 1):
            ws2.write_number(r, c, v, tot_num)
        ws2.freeze_panes(1, 1)

        ws3 = wb.add_worksheet('Detalle')
        cols = [('Fecha', 11), ('Recibo / Sesión', 18), ('Pago / Ticket', 22), ('Cliente', 36),
                ('Sucursal', 20), ('Medio', 15), ('Diario', 28), ('Tipo', 11), ('Importe', 16)]
        for c, (txt, ancho) in enumerate(cols):
            ws3.write(0, c, txt, head)
            ws3.set_column(c, c, ancho)
        for r, f in enumerate(d['detalle'], 1):
            ws3.write_datetime(r, 0, datetime.combine(f['fecha'], time.min), fecha)
            for c, k in enumerate(['recibo', 'pago', 'cliente', 'sucursal', 'medio', 'diario',
                                   'tipo'], 1):
                ws3.write(r, c, f[k])
            ws3.write_number(r, 8, f['importe'], num)
        ws3.freeze_panes(1, 0)
        ws3.autofilter(0, 0, max(len(d['detalle']), 1), len(cols) - 1)
