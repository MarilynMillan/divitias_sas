
from odoo import fields, models, api
from datetime import date

class HRPayslip(models.Model):
    _inherit = 'hr.payslip'
    
    correction_status = fields.Selection([
        ('actual', 'ACTUAL'),
        ('coreccion', 'CORRECIÓN'),
        ('no', 'NO'),
    ], string='Corrección', default='no', copy=False)