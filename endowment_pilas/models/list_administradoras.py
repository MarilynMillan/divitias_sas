# -*- coding: utf-8 -*-

from odoo import models, fields, api, _

class ListAdmin(models.Model):
    _name = 'list.administradoras'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _description = 'Lista de Administradora'

    name = fields.Char(
        string='Nombre',
        required=True,
        tracking=True
    )

    type_entity = fields.Selection([
        ('salud', 'EPS'),
        ('pension', 'AFP'),
        ('arl', 'ARL / Riesgos Laborales'),
        ('ccf', 'CCF / Caja de Compensación'),
        ('otros', 'Otros'),
    ],
        string='Tipo de Entidad',
        required=True,
        default='otros',
        tracking=True
    )

    active = fields.Boolean(default=True)