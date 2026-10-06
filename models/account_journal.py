# -*- coding: utf-8 -*-
from odoo import fields, models

TIPOS_INGRESO = [
    ('cash', 'Efectivo'),
    ('bank', 'Transferencias'),
    ('mp', 'Mercado Pago'),
    ('card', 'Tarjetas'),
    ('check', 'Cheques'),
    ('echeq', 'eCheq'),
    ('retention', 'Retenciones (no es dinero)'),
    ('other', 'Otros'),
]


class AccountJournal(models.Model):
    _inherit = 'account.journal'

    yaguven_ingreso_tipo = fields.Selection(
        TIPOS_INGRESO,
        string='Tipo en reporte de ingresos',
        help='Medio de cobro con que este diario aparece en los reportes de ingresos.\n'
             'Un diario sin tipo aparece en «Otros» para que no se pierda.\n'
             'Retenciones se informa aparte: no suma a lo que entró.',
    )
