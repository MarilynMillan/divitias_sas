
from odoo import fields, models

class HrLeaveType(models.Model):
    _inherit = 'hr.leave.type'

    # Lista de códigos de novedades PILA relevantes para Ausencias
    WORK_ENTRY_NOVELTY_CODES = [
        ('SLN', 'SLN - Suspensión del Contrato'),
        ('IGE', 'IGE - Incapacidad por Enfermedad General'),
        ('LMA', 'LMA - Licencia de Maternidad'),
        ('VAL-LR', 'VAL-LR - Vacaciones o Licencia Remunerada'),
        ('AVP', 'AVP - Ausencia por Vacaciones Pagadas'), # Aunque el Excel no lo pide, lo clasificamos
        ('VCT', 'VCT - Vacaciones Colectivas'),
        ('IRL', 'IRL - Incapacidad por Accidente de Trabajo / Enfermedad Laboral'),
        # Puedes añadir más si es necesario: TEP (Teletrabajo), etc.
    ]

    pila_novelty_code = fields.Selection(
        WORK_ENTRY_NOVELTY_CODES,
        string='Código Novedad',
        help="Clasifica esta entrada de trabajo como una novedad específica para el reporte PILA."
    )
    