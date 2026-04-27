from odoo import models, fields

class HrWorkCenter(models.Model):
    _name = 'hr.work.center'
    _description = 'Centro de Trabajo PILA'

    name = fields.Char(string="Name", required=True)
    code = fields.Char(string="Código / ID Centro")
    active = fields.Boolean(string="Activo",default=True)
    company_id = fields.Many2one('res.company', string='Compañía', default=lambda self: self.env.company)