# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import ValidationError

class administradoras(models.Model):
    _name = 'hr.administradoras' # Aseguramos el nombre técnico
    _inherit = ['hr.administradoras', 'mail.thread', 'mail.activity.mixin']


    tarifa = fields.Float( string="Tarifa",tracking=True)
    tarifa_sena = fields.Float( string="Tarifa SENA" ,tracking=True)
    tarifa_icbf = fields.Float( string="Tarifa ICBF" ,tracking=True)
    tarifa_esap = fields.Float( string="Tarifa ESAP" ,tracking=True)
    tarifa_men = fields.Float( string="Tarifa MEN" ,tracking=True)
    list_administradora_id = fields.Many2one(
        'list.administradoras',
        string="Administradora",
        domain="[('type_entity', '=', type_entity)]",
        tracking=True
    )

    traslado = fields.Boolean(
        string='Traslado',
        tracking=True,default=False)

    list_administradora_destino_id = fields.Many2one(
        'list.administradoras',
        string='Administradora Destino',
        domain="[('id', '!=', list_administradora_id)]",
        tracking=True
    )


    # --- CAMPO NUEVO ---

    type_entity = fields.Selection([
        ('salud', 'EPS'),
        ('pension', 'AFP'),
        ('arl', 'ARL / Riesgos Laborales'),
        ('ccf', 'CCF / Caja de Compensación'),
        ('otros', 'Otros')
    ], default="otros",

    string='Tipo de Entidad',
    required=True,
    help="Clasificación de la administradora para propósitos de nómina (PILA)." ,tracking=True)


    show_ccf = fields.Boolean(compute="_compute_show_ccf",store=True)
    contract_id = fields.Many2one('hr.contract', string="Contrato Relacionado")


    @api.constrains('list_administradora_id','list_administradora_destino_id','traslado')
    def _check_administradora_destino(self):
        for rec in self:
            if (
                rec.traslado
                and rec.list_administradora_id
                and rec.list_administradora_destino_id
                and rec.list_administradora_id == rec.list_administradora_destino_id
            ):
                raise ValidationError(
                    "La administradora destino debe ser diferente a la administradora origen."
                )

    @api.onchange('traslado')
    def _onchange_traslado(self):
        if not self.traslado:
            self.list_administradora_destino_id = False


    @api.depends('type_entity')
    def _compute_show_ccf(self):
        for record in self:
            record.show_ccf = record.type_entity == 'ccf'

    # MODELO hr.administradoras
    def _get_admin_label_by_type(self, entity_type):
        """
        Devuelve el nombre de la administradora SOLO
        si coincide con el tipo de entidad solicitado.
        """
        self.ensure_one()

        admin = self.list_administradora_id

        if admin and admin.type_entity == entity_type:
            return admin.name

        return ''

    def _get_admin_destino_by_type(self, entity_type):
        """
        Toma como referencia el tipo de la entidad actual (self) para
        determinar en qué columna del reporte debe aparecer el destino.
        """
        if not self:
            return ''

        self.ensure_one()

        # 1. Si no hay traslado, no mostramos nada.
        if not self.traslado:
            return ''

        # 2. Referencia: El tipo de la administradora actual (origen).
        # Si este registro es de 'salud', su destino debe ir a la celda de salud.
        if self.type_entity == entity_type:
            # Retornamos el nombre de la administradora destino
            if self.list_administradora_destino_id:
                return self.list_administradora_destino_id.name or ''

        return ''

    def get_pension_destino_label(self):
        return self._get_admin_destino_by_type('pension')

    def get_salud_destino_label(self):
        return self._get_admin_destino_by_type('salud')


    def get_pension_label(self):
        return self._get_admin_label_by_type('pension')


    def get_salud_label(self):
        return self._get_admin_label_by_type('salud')


    def get_arl_label(self):
        return self._get_admin_label_by_type('arl')


    def get_ccf_label(self):
        return self._get_admin_label_by_type('ccf')

    # Limpia las selecciones si cambias el tipo de entidad
    @api.onchange('type_entity')
    def _onchange_type_entity(self):
        self.list_administradora_id = False
        self.list_administradora_destino_id = False

    def _format_tracking_value(self, field_name, value):
        """ Convierte IDs y códigos técnicos en nombres legibles """
        if not value: return 'Vacío'
        field = self._fields[field_name]
        if field.type == 'selection':
            return dict(field.selection).get(value, value)
        if field.type == 'many2one':
            if isinstance(value, int):
                return self.env[field.comodel_name].browse(value).display_name or 'Sin Nombre'
            return value.display_name
        return str(value)

    def write(self, vals):
        # 1. Capturar cambios antes de guardar (Solo campos con tracking=True)
        updates_to_post = {}
        for record in self:
            changes = []
            for field_name, new_val in vals.items():
                field = self._fields.get(field_name)
                if field and getattr(field, 'tracking', False):
                    old_val = getattr(record, field_name)

                    old_str = record._format_tracking_value(field_name, old_val)
                    new_str = record._format_tracking_value(field_name, new_val)

                    if old_str != new_str:
                        # Implementación de la flecha con estilo
                        linea = f"""
                            <li>
                                <b>{field.string}:</b> {old_str}
                                <i class='fa fa-long-arrow-right' style='margin: 0 5px; color: #714B67;'></i>
                                {new_str}
                            </li>"""
                        changes.append(linea)

            if changes:
                updates_to_post[record.id] = changes

        # 2. Guardar en la base de datos
        res = super(administradoras, self).write(vals)

        # 3. Publicar usando contract_id (Relación directa)
        for record in self:
            if record.id in updates_to_post and record.contract_id:
                body = f"""
                    <div>
                        <b>Actualización en Administradora ({record.name}):</b>
                        <ul style='margin-top: 5px;'>{''.join(updates_to_post[record.id])}</ul>
                    </div>"""

                # Publicar mensaje directamente en el contrato vinculado
                record.contract_id.message_post(
                    body=body,
                    message_type='notification',
                    subtype_xmlid='mail.mt_note'
                )
                # Forzar refresco de la UI tocando el write_date del contrato
                record.contract_id.write({'write_date': fields.Datetime.now()})

        return res

    def create(self, vals):
        # 1. Ejecutar el create original
        record = super(administradoras, self).create(vals)

        # 2. Notificar al contrato si existe la relación contract_id
        if record.contract_id:
            tipo = record._format_tracking_value('type_entity', record.type_entity)
            tarifa = record.tarifa

            body = f"""
                <div style="background-color: #f8f9fa; padding: 10px; border-radius: 5px; border-left: 4px solid #714B67;">
                    <p style="color: #2c3e50; margin-bottom: 5px;">🌟 <b>Nueva Administradora vinculada:</b> {record.name}</p>
                    <ul style="margin: 0;">
                        <li><b>Tipo:</b> {tipo}</li>
                        <li><b>Tarifa Base:</b> {tarifa}</li>
                    </ul>
                </div>
            """
            record.contract_id.message_post(body=body, subtype_xmlid='mail.mt_note')
            # Forzar refresco inmediato
            record.contract_id.write({'write_date': fields.Datetime.now()})

        return record
