from odoo import models, fields, api, _
from odoo.exceptions import UserError


class PilaNoveltyType(models.Model):
    _name = 'pila.novelty.type'
    _description = 'Tipos de Novedades PILA'


    # Lista de códigos de novedades PILA como una tupla de tuplas
    NOVELTY_CODES = [
        ('TDE', 'TDE'),
        ('TAE', 'TAE'),
        ('TDP', 'TDP'),
        ('TAP', 'TAP'),
        # Puedes agregar más códigos PILA según los necesites (ej. IN, RT, SLN, VSP, etc.)
    ]

    name = fields.Char(string='Nombre Novedad', required=True)
    code = fields.Selection(
        NOVELTY_CODES, 
        string='Código', 
        required=True, 
        help="El código de la novedad PILA: U, V, W, TDE, TAE, etc."
    )
    
    description = fields.Text(string='Descripción')

