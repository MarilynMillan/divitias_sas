
from odoo import models, fields

class HrEmployee(models.Model):
    _inherit = "hr.employee"  
    
    
    date_resident = fields.Date(string="Fecha de Residencia")
    # Campo para el tipo de identificación (como en res.partner)
    l10n_latam_identification_type_id = fields.Many2one(
        'l10n_latam.identification.type', 
        string='Tipo de Identificación',
        help="Tipo de documento de identidad (NIT, CC, etc.)"
    )

    # Campo para el número de identificación adicional (UPC)
    upc_identification_number = fields.Char(
        string='Número de Identificación UPC',
        help="Número de identificación para el registro UPC"
    )