# -*- coding: utf-8 -*-
from odoo import models


class ReportIngresos(models.AbstractModel):
    _name = 'report.yaguven_reporte_ingresos.ingresos_doc'
    _description = 'Reporte de ingresos'

    def _get_report_values(self, docids, data=None):
        docs = self.env['yaguven.ingresos.wizard'].browse(docids)
        return {
            'doc_ids': docids,
            'doc_model': 'yaguven.ingresos.wizard',
            'docs': docs,
            # El PDF y el Excel salen del mismo cálculo.
            'datos': docs[:1]._datos(),
        }
