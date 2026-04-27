from odoo import models, fields

class PilaTipoTrabajador(models.Model):
    _name = 'pila.tipo.trabajador'
    _description = 'Tipo Trabajador'

    code = fields.Char(string='Código', required=True)
    name = fields.Char(string='Name', required=True)
    active = fields.Boolean(default=True)

    def name_get(self):
        result = []
        for rec in self:
            result.append((rec.id, f"[{rec.code}] {rec.name}"))
        return result

class PilaSubtipoTrabajador(models.Model):
    _name = 'pila.subtipo.trabajador'
    _description = 'Subtipo Trabajador'

    code = fields.Char(string='Código', required=True)
    name = fields.Char(string='Name', required=True)
    active = fields.Boolean(default=True)

    def name_get(self):
        result = []
        for rec in self:
            result.append((rec.id, f"[{rec.code}] {rec.name}"))
        return result
