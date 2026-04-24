from odoo import models, fields, api, _
from odoo.exceptions import UserError
from datetime import date


class HrContractNoveltyLine(models.Model):
    _name = 'hr.contract.novelty.line'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _description = 'Novedades aplicadas al Contrato'
    _order = 'date_start desc, novelty_type_id'

    contract_id = fields.Many2one('hr.contract', string='Contrato', required=True, ondelete='cascade')
    novelty_type_id = fields.Many2one(
        'pila.novelty.type', 
        string='Tipo de Novedad', 
        required=True,
        help="Selecciona la novedad configurada (TAE, TDP, TAP, etc.)"
    )

    date_start = fields.Date(string='Fecha Inicio', required=True)
    date_end = fields.Date(string='Fecha Fin')

    def write(self, vals):
        tracked_fields = {
            'novelty_type_id': 'Tipo de Novedad',
            'date_start': 'Fecha Inicio',
            'date_end': 'Fecha Fin',
        }

        changes_dict = {}

        for record in self:
            changes = []
            
            # --- 1. Capturar Información Antes del Write ---
            old_novelty_name = record.novelty_type_id.display_name if record.novelty_type_id else 'Sin Tipo'
            
            for field_name, field_label in tracked_fields.items():
                if field_name in vals:
                    old_value = getattr(record, field_name)
                    new_value_raw = vals.get(field_name)

                    # Formatea valores (usa la función auxiliar que maneja las fechas y M2O)
                    old_str = record._format_tracking_value(field_name, old_value)
                    new_str = record._format_tracking_value(field_name, new_value_raw)

                    if old_str != new_str:
                        # Formato: [Etiqueta] de [Valor Antiguo] a [Valor Nuevo]
                        # Usamos <strong> para la etiqueta y ** para los valores
                        changes.append(
                            f"{field_label}: {old_str} → {new_str}"
                        )
            
            if changes:
                # Almacenamos los cambios y el nombre anterior de la novedad
                changes_dict[record.id] = {'changes': changes, 'old_name': old_novelty_name}

        # 2. Ejecutar el write original para aplicar los cambios
        res = super().write(vals)

        # 3. Publicar cambios en el contrato
        for record in self:
            data = changes_dict.get(record.id)
            if data and record.contract_id:
                record_changes = data['changes']
                old_name = data['old_name']
                new_name = record.novelty_type_id.display_name if record.novelty_type_id else 'Sin Tipo'
                
                # --- Generación del Encabezado (Más Descriptivo) ---
                if 'novelty_type_id' in vals and old_name != new_name:
                    # El tipo de novedad fue lo que cambió
                    header = f"El tipo de novedad cambió de **{old_name}** a **{new_name}**."
                else:
                    # Solo cambiaron las fechas u otros campos
                    header = f"Se actualizaron datos en la novedad **{new_name}**."
                
                # Juntamos el encabezado y el detalle de los cambios con formato HTML <ul>
                body = f"{header}" + "</li><li>".join(record_changes) + "</li></ul>"
                
                record.contract_id.message_post(body=body)

        return res

    # ----------------------------
    # Create: mensaje al crear
    # ----------------------------
    def create(self, vals):
        record = super().create(vals)
        if record.contract_id:
            novelty_name = record.novelty_type_id.display_name if record.novelty_type_id else 'Sin Tipo'
            date_start = record.date_start.strftime('%Y-%m-%d') if record.date_start else 'No asignada'
            date_end = record.date_end.strftime('%Y-%m-%d') if record.date_end else 'No asignada'

            body = (
                f"Se creó una nueva novedad en el contrato: **{novelty_name}**.("
                f"Fecha Inicio: {date_start}-"
                f"Fecha Fin: {date_end})"
            )
            record.contract_id.message_post(body=body)
        return record

    # ----------------------------
    # Unlink: mensaje al eliminar
    # ----------------------------
    def unlink(self):
        for record in self:
            if record.contract_id:
                novelty_name = record.novelty_type_id.display_name if record.novelty_type_id else 'Sin Tipo'
                body = f"Se eliminó la novedad **{novelty_name}** del contrato."
                record.contract_id.message_post(body=body)
        return super().unlink()

    # ----------------------------

    def _format_tracking_value(self, field_name, value):
        """Formatea el valor para mostrarlo en el historial, manejando la conversión de tipos (str a date)."""
        field = self._fields.get(field_name)

        if not value:
            return 'No asignado'
        
        # --- Manejo de Fechas (CORRECCIÓN CLAVE) ---
        if field.type in ('date', 'datetime'):
            # Si el valor es una cadena (viene de 'vals' en el write), la convertimos al objeto de fecha/hora.
            if isinstance(value, str):
                if field.type == 'date':
                    # fields.Date.from_string convierte 'YYYY-MM-DD' string a objeto date
                    value = fields.Date.from_string(value)
                elif field.type == 'datetime':
                    # fields.Datetime.from_string convierte string a objeto datetime
                    value = fields.Datetime.from_string(value)
            
            # Ahora que 'value' es un objeto date/datetime, usamos el método de formato de Odoo.
            if field.type == 'date':
                return fields.Date.to_string(value)
            elif field.type == 'datetime':
                return fields.Datetime.to_string(value)

        # --- Manejo de Many2one ---
        elif field.type in ('many2one', 'many2many'):
            # Si el valor es un ID (int), buscamos el nombre (puede ocurrir si se pasa un ID en vals)
            if isinstance(value, int):
                if value == 0:
                    return 'No asignado'
                related_model = self.env[field.comodel_name]
                related_record = related_model.browse(value)
                return related_record.display_name if related_record.exists() else str(value)

            # Si el valor es un recordset (old_value o después de browse), usa display_name
            return value.display_name if value else 'No asignado'
        
        # Para todos los demás tipos (Char, Float, Selection, etc.)
        return str(value)