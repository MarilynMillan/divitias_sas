from odoo import models, fields

class HrContractConcepto(models.Model):
    _name = 'hr.contract.concept'
    _description = 'Conceptos de contrato'

    name = fields.Char(string="Name", required=True)
    type = fields.Selection([
        ('ingreso', 'Ingreso'),
        ('retiro', 'Egreso')
    ], string="Tipo de Concepto", required=True)
