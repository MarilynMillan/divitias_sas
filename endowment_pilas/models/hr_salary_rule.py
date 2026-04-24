
from odoo import models, fields, api, _

class HrSalaryRule(models.Model):
    _inherit = 'hr.salary.rule'
  
    is_payment_transitory = fields.Boolean(string="Pago extra")


    tipo_reporte_excel = fields.Selection([
        ('valor_cotizacion_pension', 'Valor Cotización Pensión'),
        ('valor_cotizacion_salud', 'Valor Cotización Salud'),
        ('valor_cotizacion_riesgo', 'Valor Cotización Riesgos (ARL)'),
        ('valor_cotizacion_ccf', 'Valor Cotización CCF'),
        ('cotizacion_voluntaria_afiliado', 'Cotización Voluntaria Afiliado'),
        ('cotizacion_voluntaria_empleador', 'Cotización Voluntaria Empleador'),
        ('fondo_solidaridad', 'Fondo Solidaridad'),
        ('fondo_subsistencia', 'Fondo Subsistencia'),
        ('valor_no_retenido', 'Valor No Retenido'),
        ('total_aportes', 'Total Aportes'),
        ('valor_upc', 'Valor UPC'),
        ('valor_incapacidad_eg', 'Valor Incapacidad EG'),
        ('valor_licencia_maternidad', 'Valor Licencia Maternidad'),
        ('ibc', 'IBC'),
        ('ibc_otros_parafiscales', 'IBC Otros Parafiscales'),
        ('valor_cotizacion_sena', 'Sena'),
        ('valor_cotizacion_icbf', 'Cotización ICBF'),
        ('valor_cotizacion_esap', 'Cotización ESAP'),
        ('valor_cotizacion_men', 'Cotización MEN'),
        ('exonerado_1607', 'Exonerado Ley 1607'),
    ], string="Tipo de Regla para Reporte Excel", help="Identifica esta regla para el reporte de seguridad social")

