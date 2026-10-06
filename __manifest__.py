# -*- coding: utf-8 -*-
{
    "name": "Reportes de Ingresos (Yagüven)",
    "summary": "Lo que entró por cada medio de cobro: reporte diario e ingresos por período, "
               "en PDF y Excel, por sucursal.",
    "description": """
Port a Odoo 20 de los reportes de ingresos de Lupatini en Odoo 17 (lupatini_reporte_diario
17.0.1.4.5), con el criterio que se corrigió ahí:

- Sólo cobros de CLIENTES: los pagos a proveedores hechos por la misma caja no restan.
- Las reversiones (pagos salientes a clientes, NC devueltas en plata) restan del medio.
- Las retenciones que nos hacen los clientes se informan aparte: no son plata.
- El medio de cobro sale del «Tipo en reporte de ingresos» del diario, configurable desde
  la UI. Un diario sin tipo aparece en «Otros» para que su plata no desaparezca del total.

Diferencia con el 17: en Odoo 20 el mostrador cobra con pagos de POS. Se suman los pagos de
POS con diario (por el diario de su método) y se excluyen los pagos contables que Odoo crea a
partir de ellos (`pos_payment_method_id` con diario), que si no se contarían dos veces. Los
cheques y retenciones del mostrador (yaguven_pos_cheque) entran por su pago contable.
""",
    "version": "20.0.1.0.2",
    "category": "Accounting",
    "author": "Yagüven C.G.",
    "license": "LGPL-3",
    "depends": [
        "account",
        "point_of_sale",
        "yaguven_operating_unit",
        "yaguven_payment_group",
    ],
    "data": [
        "security/ir.access.csv",
        "views/account_journal_views.xml",
        "wizard/ingresos_wizard_views.xml",
        "report/report_ingresos_templates.xml",
    ],
    "external_dependencies": {"python": ["xlsxwriter"]},
    "installable": True,
    "application": False,
}
