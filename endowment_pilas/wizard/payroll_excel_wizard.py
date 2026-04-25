import pandas as pd
import io
import os
import base64
import openpyxl
from odoo import fields, models, api, _
from odoo.exceptions import UserError
from odoo.modules import get_module_resource
from openpyxl.styles import Alignment
from datetime import datetime, date
from copy import copy

import logging

_logger = logging.getLogger(__name__)




class PayrollExcelWizard(models.TransientModel):
    _name = 'hr.payroll.excel.wizard'
    _description = 'Wizard para Reporte de Nómina a Excel'

    date_from = fields.Date(string="Fecha Desde")
    date_to = fields.Date(string="Fecha Hasta")
    payslip_run_id = fields.Many2one('hr.payslip.run', string="Lote de Nómina")

    plantilla_excel = fields.Binary(string='Plantilla Excel', help='Sube la plantilla de Excel para usar en el reporte. Si está vacío, se usará la plantilla por defecto.')
    plantilla_excel_name = fields.Char(string='Nombre Archivo Plantilla')

    #file_input = fields.Binary(string="Archivo Excel del Cliente", required=True)
    #file_name = fields.Char(string="Nombre del Archivo")



    def action_generate_excel_report(self):


        """
        Genera el reporte Excel de recibos de nómina filtrados.
        """
        # 1. DEFINICIÓN DEL DOMINIO Y BÚSQUEDA
        domain = [('state', '=', 'done')]
        if self.date_from: domain.append(('date_from', '>=', self.date_from))
        if self.date_to: domain.append(('date_to', '<=', self.date_to))
        if self.payslip_run_id: domain.append(('payslip_run_id', '=', self.payslip_run_id.id))

        payslips = self.env['hr.payslip'].search(domain)
        if not payslips:
            raise UserError(_("Nessun cedolino trovato per i filtri selezionati."))

        # 2. CARICAMENTO MODELLO (Ottimizzato per non corrompere il file)
        output = io.BytesIO()

        if self.plantilla_excel:
            # Opzione A: File caricato dall'utente
            input_buffer = io.BytesIO(base64.b64decode(self.plantilla_excel))
            wb = openpyxl.load_workbook(input_buffer, data_only=False)
        else:
            # Opzione B: File "DIVITIASSAS.xlsx" nella cartella data
            file_name_in_data = 'DIVITIASSAS.xlsx'
            path = get_module_resource('endowment_pilas', 'data', file_name_in_data)
            if not path or not os.path.exists(path):
                raise UserError(_("El modelo DIVITIASSAS.xlsx no se encuentra en la hoja de datos del módulo."))
            wb = openpyxl.load_workbook(path, data_only=False)

        # Cerchiamo il foglio "Liquidaciones" o quello attivo
        try:
            sheet = wb['Liquidaciones']
        except KeyError:
            sheet = wb.active

        # 3. LLENADO DE DATOS (Empezamos en la fila 19 porque la 17 y 18 son encabezados)
        row_num = 19
        row_counter = 1

        for payslip in payslips:
            employee = payslip.employee_id
            contract = payslip.contract_id
            correction_status_value = payslip.correction_status if hasattr(payslip, 'correction_status') else 'No'


            ############################## type identification#########################
            # Obtenemos el registro del tipo de identificación

            mapeo_id = {
                'Cédula de ciudadanía': 'CC',
                'Cédula de extranjería': 'CE',
                'Tarjeta de Identidad': 'TI',
                'Registro Civil': 'RC',
                'Pasaporte': 'PA',
                'Permiso por Protección Temporal': 'PT', # El que viste en la lista
                'PEP (Permiso Especial de Permanencia)': 'PE',
                'NIT': 'NI',
                'ID Extranjera': 'CE',
                'Documento de identificación extranjero': 'CD'
            }

            # 1. Obtenemos el registro
            tipo_doc_rec = employee.employee_address_home.l10n_latam_identification_type_id

            # 2. Extraemos el nombre (string)
            nombre_largo = tipo_doc_rec.name or ''

            # 3. Aplicamos el mapeo que definimos antes para que salga "CC", "CE", etc.
            # mapeo_id es el diccionario que definimos en el paso anterior
            tipo_doc_abreviado = mapeo_id.get(nombre_largo, nombre_largo)
            #####################################################################
            tipo_cotizante_excel_label = 'NINGUNA'
            if contract:
                if contract.pila_tipo_trabajador_id:
                    # Usamos el nombre del nuevo campo configurable, si tiene algo (como '1.Dependiente')
                    tipo_cotizante_excel_label = contract.pila_tipo_trabajador_id.name or "NINGUNA"
                #elif contract.tipo_trabajador:
                    # --- CORRECCIÓN AQUÍ ---
                    # Método recomendado para obtener la etiqueta legible del campo Selection
                    #tipo_cotizante_excel_label = dict(contract._fields['tipo_trabajador'].selection).get(contract.tipo_trabajador, '')

            sub_cotizante_excel_label = 'NINGUNA'
            if contract:
                if contract.pila_subtipo_trabajador_id:
                    sub_cotizante_excel_label = contract.pila_subtipo_trabajador_id.name or "NINGUNA"
                #elif contract.sub_tipo_trabajador:
                    # --- CORRECCIÓN AQUÍ ---
                    # Método recomendado para obtener la etiqueta legible del campo Selection
                    #sub_cotizante_excel_label = dict(contract._fields['sub_tipo_trabajador'].selection).get(contract.sub_tipo_trabajador, '')

            # --- NUEVO: Lógica para 'Horas Laboradas' ---

            dias_cotizados_pension = 0.0
            dias_cotizados_salud = 0.0
            dias_cotizados_arl = 0.0
            dias_cotizados_ccf = 0.0

            horas_laboradas = 0.0
            dias_laborados = 0.0  # 👈 nuevo acumulador

            for worked_day_line in payslip.worked_days_line_ids:
                days = worked_day_line.number_of_days
                hours = worked_day_line.number_of_hours
                code = worked_day_line.code

                # DÍAS: Se acumulan todos para que coincida con el total (sum) del XML
                dias_laborados += days

                # HORAS: Únicamente si es el código de asistencia
                if code == 'WORK100':
                    horas_laboradas += hours


                # 2. ACUMULACIÓN DE DÍAS COTIZADOS
                # Se utiliza el código de la línea del día trabajado directamente
                # para determinar a qué columna PILA se suma.

                if code == 'pension':
                    dias_cotizados_pension += days
                elif code == 'salud':
                    dias_cotizados_salud += days
                elif code == 'arl':
                    dias_cotizados_arl += days
                elif code == 'ccf':
                    dias_cotizados_ccf += days


            ####################################################

            # Inicializa ambas etiquetas
            es_extranjero_label = 'No'
            es_residente_label = 'No'

            # Lógica para 'Extranjero' (Basada en País)
            if employee and employee.country_id:
                if employee.country_id.name != 'Colombia':
                    es_extranjero_label = 'Si'

            if employee and employee.is_non_resident:
                es_residente_label = 'Si'



            # --- Lógica: Fecha de radicación en el exterior ---
            fecha_radicacion = ''
            if employee and getattr(employee, 'date_resident', False):
                fecha_radicacion = employee.date_resident  # Debe ser tipo Date en tu modelo

             # --- Lógica: Fecha inicio del contrato ---
            fecha_inicio_contrato = ''
            if contract and getattr(contract, 'date_start', False):
                fecha_inicio_contrato = contract.date_start  # Tipo Date normalmente


            # --- Lógica: ING (Ingreso) ---
            valor_ing = 'NO'
            if fecha_inicio_contrato:
                es_periodo_ingreso = False
                # Validación de rango
                if self.date_from and self.date_to:
                    if self.date_from <= fecha_inicio_contrato <= self.date_to:
                        es_periodo_ingreso = True
                elif self.date_from:
                    if fecha_inicio_contrato >= self.date_from:
                        es_periodo_ingreso = True
                elif self.date_to:
                    if fecha_inicio_contrato <= self.date_to:
                        es_periodo_ingreso = True
                else:
                    es_periodo_ingreso = True

                # Si está en el rango, aplicamos el concepto del contrato
                if es_periodo_ingreso:
                    # Buscamos el código del Many2one, si no hay, ponemos 'X'
                    valor_ing = contract.pila_ingreso_concepto_id.name or 'NO'

            # --- Lógica: Fecha final del contrato ---
            fecha_final_contrato = False
            if contract and getattr(contract, 'date_end', False):
                fecha_final_contrato = contract.date_end

            # --- Lógica: RET (Retiro) ---
            valor_ret = 'NO'
            if fecha_final_contrato:
                es_periodo_retiro = False
                # Validación de rango
                if self.date_from and self.date_to:
                    if self.date_from <= fecha_final_contrato <= self.date_to:
                        es_periodo_retiro = True
                elif self.date_from:
                    if fecha_final_contrato >= self.date_from:
                        es_periodo_retiro = True
                elif self.date_to:
                    if fecha_final_contrato <= self.date_to:
                        es_periodo_retiro = True
                else:
                    es_periodo_retiro = True

                # Si está en el rango, aplicamos el concepto del contrato
                if es_periodo_retiro:
                    # Buscamos el código del Many2one, si no hay, ponemos 'X'
                    valor_ret = contract.pila_retiro_concepto_id.name or 'NO'

            # Inicializar un diccionario para guardar los valores por defecto 'NO'
            # Asegúrate de que los códigos aquí coincidan con los de tu campo 'novelty_code'
            # 1. Inicializamos todas las novedades simples en 'NO' por defecto
            novelty_values = {code: 'NO' for code in ['TDE', 'TAE', 'TDP', 'TAP']} # Agrega aquí tus códigos

            report_start_date = self.date_from
            report_end_date = self.date_to

            afp_destino = ""
            eps_destino = ""

            # Supongamos que recorremos las administradoras configuradas en el contrato
            # o donde las tengas relacionadas (ajusta 'contract.administradora_ids' según tu modelo)
            if contract.administradoras_ids:
                for admin_line in contract.administradoras_ids:
                    # 1. ¿Es una administradora de tipo Pensión y tiene Traslado marcado?
                    if admin_line.type_entity == 'pension' and admin_line.traslado:
                        if admin_line.list_administradora_destino_id:
                            # Tomamos el nombre de la entidad destino de la lista maestra
                            afp_destino = admin_line.list_administradora_destino_id.name.upper()

                    # 2. ¿Es una administradora de tipo Salud y tiene Traslado marcado?
                    elif admin_line.type_entity == 'salud' and admin_line.traslado:
                        if admin_line.list_administradora_destino_id:
                            # Tomamos el nombre de la entidad destino de la lista maestra
                            eps_destino = admin_line.list_administradora_destino_id.name.upper()

            # Inicializar VSP con valores vacíos/por defecto
            valor_vsp = 'NO'
            fecha_vsp_str = ''

            # ----------------------------------------------------
            # 🆕 Lógica Simplificada para VSP (Variación Salario)
            # ----------------------------------------------------

            # 1. Obtenemos la fecha de cambio del campo dedicado del contrato
            fecha_cambio_sueldo = contract.date_wage_change

            if fecha_cambio_sueldo and report_start_date and report_end_date:

                # 2. Verificamos si la fecha de cambio cae DENTRO del rango del reporte.
                # El filtro que usa el reporte: [self.date_from, self.date_to]
                if report_start_date <= fecha_cambio_sueldo <= report_end_date:

                    # Si la fecha de cambio está dentro del periodo, es VSP = 'SI'
                    valor_vsp = 'SI'
                    fecha_vsp_str = fecha_cambio_sueldo # Ya es un objeto date

            #####################################################################
            # --- Lógica: VST (Variación Transitoria de Salario) ---
            valor_vst = 'NO'

            # Buscamos los payslips del empleado que estén dentro del rango y realizados (state = 'done')
            # 'contract' es el objeto del contrato actual en tu bucle
            payslips = self.env['hr.payslip'].search([
                ('employee_id', '=', contract.employee_id.id),
                ('date_from', '>=', report_start_date),
                ('date_to', '<=', report_end_date),
                ('state', '=', 'done')
            ])

            if payslips:
                # Revisamos las líneas de esas nóminas (slip_ids.line_ids)
                # Buscamos si alguna línea pertenece a una regla con is_payment_transitory = True
                # Y que el monto (total) sea mayor a cero
                transitory_lines = payslips.mapped('line_ids').filtered(
                    lambda l: l.salary_rule_id.is_payment_transitory and l.total > 0
                )

                if transitory_lines:
                    valor_vst = 'SI'

            # ----------------------------------------------------
            # 📌 Lógica para Ausencias (Usando hr.leave - Solicitudes de Ausencia)
            # ----------------------------------------------------

            # Definición de códigos relevantes y Inicialización a 'NO'
            ABSENCE_CODES = ['SLN', 'IGE', 'LMA', 'VAL-LR', 'AVP', 'VCT', 'IRL']

            # Tu búsqueda actual (sin cambios)
            leaves = self.env['hr.leave'].search([
                ('employee_id', '=', employee.id),
                ('state', '=', 'validate'),
                ('date_from', '<=', report_end_date),
                ('date_to', '>=', report_start_date),
            ])

            # Filtramos solo las que nos interesan
            relevant_leaves = []
            if leaves:
                for leave in leaves:
                    novelty_code = leave.holiday_status_id.pila_novelty_code
                    if novelty_code and novelty_code in ABSENCE_CODES:
                        relevant_leaves.append(leave)

            # Si no hay hojas relevantes, añadimos un None para que el bucle haga una iteración con todo "NO"
            if not relevant_leaves:
                relevant_leaves = [None]


            for current_leave in relevant_leaves:
                absence_novelties = {code: {'value': 'NO', 'start': '', 'end': ''} for code in ABSENCE_CODES}
                count_lma = 0
                count_ige = 0

                if current_leave:
                    novelty_code = current_leave.holiday_status_id.pila_novelty_code
                    if novelty_code == 'LMA':
                        count_lma = int(current_leave.number_of_days)
                    elif novelty_code == 'IGE':
                        count_ige = int(current_leave.number_of_days)

                    absence_novelties[novelty_code]['value'] = current_leave.holiday_status_id.name or 'SI'

                    if novelty_code != 'AVP':
                        absence_novelties[novelty_code]['start'] = current_leave.date_from.date()
                        absence_novelties[novelty_code]['end'] = current_leave.date_to.date()

                # Usaremos el nombre (label) para el Excel.
                # El método _get_selection_label() toma el nombre legible del campo de selección.
                correction_status_excel_value = payslip.correction_status if hasattr(payslip, 'correction_status') else 'no'

                # 2. Salario Mensual (Del campo 'wage' del contrato)
                salario_mensual = contract.wage if contract else 0.0

                        # ----------------------------------------------------
                # 📌 LÓGICA PARA SALARIO INTEGRAL Y VARIABLE
                # ----------------------------------------------------

                # 1. Salario Integral (True/False del contrato -> SI/NO)
                if contract and contract.wage_integral:
                    integral_excel_value = 'SI'
                else:
                    integral_excel_value = 'NO'

                # 2. Salario Variable (True/False del contrato -> SI/NO)
                if contract and contract.wage_variable:
                    variable_excel_value = 'SI'
                else:
                    variable_excel_value = 'NO'

                ##################################################################

                # ADMINISTRADORAS
                pension_admin = contract.get_admin_by_type('pension')
                salud_admin = contract.get_admin_by_type('salud')
                arl_admin = contract.get_admin_by_type('arl')
                ccf_admin = contract.get_admin_by_type('ccf')


                # --------------------------- Lógica de Alto Riesgo ---------------------------
                alto_riesgo_selection = self.env['hr.contract'].fields_get(allfields=['alto_riesgo']
                )['alto_riesgo']['selection']

                alto_riesgo_label = ''
                if contract and contract.alto_riesgo:
                    alto_riesgo_label = dict(alto_riesgo_selection).get(contract.alto_riesgo, '')


                claves_reporte = [
                    'valor_cotizacion_pension', 'valor_cotizacion_salud', 'valor_cotizacion_riesgo',
                    'valor_cotizacion_ccf', 'cotizacion_voluntaria_afiliado', 'cotizacion_voluntaria_empleador',
                    'fondo_solidaridad', 'fondo_subsistencia', 'valor_no_retenido', 'total_aportes',
                    'valor_upc', 'valor_incapacidad_eg', 'valor_licencia_maternidad',
                    'ibc','ibc_otros_parafiscales',
                    'valor_cotizacion_sena', 'valor_cotizacion_icbf', 'valor_cotizacion_esap',
                    'valor_cotizacion_men', 'exonerado_1607'
                ]

                # 2. Inicializar con 0.0 (es mejor para cálculos numéricos)
                valores_reglas = {k: 0.0 for k in claves_reporte}

                # 3. Recorrer las líneas de la nómina
                for line in payslip.line_ids:
                    # Obtenemos la marca de la regla
                    tipo = line.salary_rule_id.tipo_reporte_excel

                    # Si la regla tiene una marca y esa marca está en nuestras claves
                    if tipo and tipo in valores_reglas:
                        valores_reglas[tipo] += line.total

                ######################Otras tarifas ##################################
                admins = payslip.contract_id.administradoras_ids
                #t_ccf = sum(admins.mapped('tarifa_ccf')) or 0.0
                t_sena = sum(admins.mapped('tarifa_sena')) or 0.0
                t_icbf = sum(admins.mapped('tarifa_icbf')) or 0.0
                t_esap = sum(admins.mapped('tarifa_esap')) or 0.0
                t_men = sum(admins.mapped('tarifa_men')) or 0.0

                ################Clase ################################

                valor_clase = contract.clase if contract.clase else ''

                ################### centro trabajo ################################

                #nombre_departamento = contract.department_id.name if contract.department_id else ''
                val_centro_trabajo = ""
                if contract.work_center_id:
                    val_centro_trabajo = contract.work_center_id.name.upper()
                else:
                    # SI NO HAY REGISTRO, PASAMOS EL VALOR FIJO
                    val_centro_trabajo = "RIESGO III"

                #####################Actividad Economica #######################################
                actividad_economica = contract.economic_activitity if contract.economic_activitity else ''

                ################################upc adicional################################
                mapeo_id = {
                    'Cédula de ciudadanía': 'CC',
                    'Tarjeta de identidad': 'TI',
                    'Registro civil': 'RC',
                    'Cédula de extranjería': 'CE',
                    'Pasaporte': 'PA',
                    'NIT': 'NI',
                    'Permiso Especial de Permanencia': 'PE',
                    'Permiso por Protección Temporal': 'PT',
                }

                tipo_doc_upc_rec = employee.l10n_latam_identification_type_id

                # 2. Extraemos el nombre (o cadena vacía si no hay)
                nombre_doc_upc = tipo_doc_upc_rec.name or ''

                # 3. Buscamos la abreviatura en el mapeo. Si no está, dejamos el nombre original.
                tipo_doc_upc_abreviado = mapeo_id.get(nombre_doc_upc, nombre_doc_upc)

                # 4. Obtenemos el número y le quitamos puntos/guiones/espacios
                raw_upc = employee.upc_identification_number or ''
                numero_upc_limpio = str(raw_upc).replace('.', '').replace('-', '').strip()
                # 4. Obtenemos el número de identificación adicional
                #numero_upc = employee.upc_identification_number or ''

                data = [
                    row_counter,                                      # 1 (A)
                    tipo_doc_abreviado or 'CC',                                # 2 (B) - Debe ser 'CC', 'CE', etc.
                    str(employee.employee_address_home.vat or '').replace('.', '').replace('-', '').strip(), # 3 (C)
                    (employee.employee_address_home.last_name or '').upper(), # 4 (D)
                    (employee.employee_address_home.second_last_name or '').upper(), # 5 (E)
                    (employee.employee_address_home.first_name or '').upper(), # 6 (F)
                    (employee.employee_address_home.middle_name or '').upper(), # 7 (G)
                    (employee.employee_address_home.state_id.name or 'BOGOTA').upper(), # 8 (H)
                    (employee.employee_address_home.city_id.name or 'BOGOTA').upper(), # 9 (I)
                    tipo_cotizante_excel_label or '1. DEPENDIENTE',   # 10 (J)
                    sub_cotizante_excel_label or 'NINGUNO',           # 11 (K)
                    horas_laboradas or 0,                             # 12 (L)
                    es_extranjero_label or 'NO',                      # 13 (M)
                    es_residente_label or 'NO',                       # 14 (N)
                    fecha_radicacion or '',                           # 15 (O)

                    # --- BLOQUE DE MARCAS (SI/NO) ---
                    valor_ing or 'NO',                                # 16 (P) ING
                    fecha_inicio_contrato or '',                      # 17 (Q) Fecha ING
                    valor_ret or 'NO',                                # 18 (R) RET
                    fecha_final_contrato or '',                       # 19 (S) Fecha RET
                    novelty_values.get('TDE', 'NO'),                  # 20 (T)
                    novelty_values.get('TAE', 'NO'),                  # 21 (U)
                    novelty_values.get('TDP', 'NO'),                  # 22 (V)
                    novelty_values.get('TAP', 'NO'),                  # 23 (W)
                    valor_vsp or 'NO',                                # 24 (X) VSP
                    fecha_vsp_str or '',                              # 25 (Y) Fecha VSP
                    'NO',                                             # 26 (Z) VST (Variación Transitoria)

                    # --- BLOQUE AUSENCIAS (3 columnas por cada una) ---
                    absence_novelties['SLN']['value'], absence_novelties['SLN']['start'], absence_novelties['SLN']['end'], # 27,28,29
                    absence_novelties['IGE']['value'], absence_novelties['IGE']['start'], absence_novelties['IGE']['end'], # 30,31,32
                    absence_novelties['LMA']['value'], absence_novelties['LMA']['start'], absence_novelties['LMA']['end'], # 33,34,35
                    absence_novelties['VAL-LR']['value'], absence_novelties['VAL-LR']['start'], absence_novelties['VAL-LR']['end'], # 36,37,38
                    absence_novelties['AVP']['value'],                # 39 (Solo marca)
                    absence_novelties['VCT']['value'], absence_novelties['VCT']['start'], absence_novelties['VCT']['end'], # 40,41,42
                    absence_novelties['IRL']['value'], absence_novelties['IRL']['start'], absence_novelties['IRL']['end'], # 43,44,45

                    correction_status_excel_value or 'NO',            # 46
                    salario_mensual or 0,                             # 47
                    integral_excel_value or 'NO',                     # 48
                    variable_excel_value or 'NO',                     # 49

                    # --- PENSION ---
                    pension_admin.get_pension_label() if pension_admin else 'NINGUNO', # 50
                    dias_laborados or 0,                      # 51
                    valores_reglas['ibc'] or 0,               # 52
                    contract.get_tarifa_by_type('pension') or 0,      # 53
                    valores_reglas['valor_cotizacion_pension'] or 0,  # 54
                    alto_riesgo_label or 'NO',                        # 55
                    valores_reglas['cotizacion_voluntaria_afiliado'] or 0, # 56
                    valores_reglas['cotizacion_voluntaria_empleador'] or 0,# 57
                    valores_reglas['fondo_solidaridad'] or 0,         # 58
                    valores_reglas['fondo_subsistencia'] or 0,        # 59
                    valores_reglas['valor_no_retenido'] or 0,         # 60
                    valores_reglas['total_aportes'] or 0,             # 61
                    pension_admin.get_pension_destino_label() if pension_admin else 'NINGUNO', # 62

                    # --- SALUD ---
                    salud_admin.get_salud_label() if salud_admin else 'NINGUNO', # 63
                    dias_laborados or 0,                        # 64
                    valores_reglas['ibc'] or 0,                 # 65
                    contract.get_tarifa_by_type('salud') or 0,        # 66
                    valores_reglas['valor_cotizacion_salud'] or 0,    # 67
                    valores_reglas['valor_upc'] or 0,                 # 68
                    count_ige or '',                                  # 69 N° Autorización EG
                    valores_reglas['valor_incapacidad_eg'] or 0,      # 70
                    count_lma or '',                                  # 71 N° Autorización LMA
                    valores_reglas['valor_licencia_maternidad'] or 0, # 72
                    salud_admin.get_salud_destino_label() if salud_admin else 'NINGUNO', # 73

                    # --- RIESGOS (ARL) ---
                    arl_admin.get_arl_label() if arl_admin else 'NINGUNO', # 74
                    dias_laborados or 0,                          # 75
                    valores_reglas['ibc'] or 0,                # 76
                    contract.get_tarifa_by_type('arl') or 0,          # 77
                    valor_clase or '1',                               # 78
                    #nombre_departamento or '',                        # 79
                    val_centro_trabajo or '',                        # 79
                    actividad_economica or '',                        # 80
                    valores_reglas['valor_cotizacion_riesgo'] or 0,   # 81

                    # --- CAJA Y PARAFISCALES ---
                    dias_laborados or 0,                          # 82
                    ccf_admin.get_ccf_label() if ccf_admin else 'NINGUNO', # 83
                    valores_reglas['ibc'] or 0,                   # 84
                    contract.get_tarifa_by_type('ccf') or 0,          # 85
                    valores_reglas['valor_cotizacion_ccf'] or 0,      # 86
                    valores_reglas['ibc_otros_parafiscales'] or 0,    # 87
                    t_sena or 0,                                      # 88
                    valores_reglas['valor_cotizacion_sena'] or 0,     # 89
                    t_icbf or 0,                                      # 90
                    valores_reglas['valor_cotizacion_icbf'] or 0,     # 91
                    t_esap or 0,                                      # 92
                    valores_reglas['valor_cotizacion_esap'] or 0,     # 93
                    t_men or 0,                                       # 94
                    valores_reglas['valor_cotizacion_men'] or 0,      # 95
                    valores_reglas['exonerado_1607'] or 0,         # 96
                    tipo_doc_upc_abreviado if numero_upc_limpio else '', # 97: Solo pon tipo si hay número
                    numero_upc_limpio,
                ]

                def clean_text(val):
                    v = str(val or '').strip().upper()
                    return "" if v in ['', 'NONE', 'FALSE'] else v


                def clean_id(val):
                    v = str(val or '').replace('.', '').replace('-', '').strip()
                    return "" if v.upper() in ['', 'NONE', 'FALSE', 'NINGUNO'] else v

                def limit(val, max_len):
                    return val[:max_len] if val else ""

                # --- PROCESO DE ESCRITURA CON VALIDACIÓN DE FORMATO ---
                for col_num, cell_value in enumerate(data, start=1):
                    cell = sheet.cell(row=row_num, column=col_num)

                    # 🔍 DEBUG DE LONGITUD (AQUÍ)
                    if len(str(cell_value or '')) > 30:
                        _logger.warning(f"⚠️ FILA {row_num} COL {col_num} VALOR LARGO: {cell_value}")

                    # ---------------- COLUMNA 1 ----------------
                    # ---------------- COLUMNA 1 (A) ----------------
                    if col_num == 1:
                        cell.value = row_counter
                        cell.number_format = '0'
                        cell.alignment = Alignment(horizontal='right', vertical='center')

                        # AJUSTE DE ANCHO:
                        # 'A' es la columna 1. Un ancho de 5 a 8 es suficiente para el contador.
                        # COMENTADO PORQUE CAUSA ERROR EN APORTES EN LINEA AL CREAR COLUMNAS METADATA
                        # sheet.column_dimensions['A'].width = 8
                        continue

                    # ---------------- IDs ----------------
                    if col_num in [2, 3, 97, 98]:
                        cell.value = clean_id(cell_value)
                        cell.data_type = 's'
                        continue

                    # ---------------- NOMBRES ----------------
                    if col_num in [4, 5, 6, 7]:
                        #cell.value = clean_text(cell_value)
                        val = clean_text(cell_value)
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                        cell.value = limit(val, 20)  # ajusta según PILA real
                        continue

                    # ---------------- ADMINISTRADORAS ----------------
                    if col_num in [50, 62, 63, 73, 74, 83]:
                        val = clean_text(cell_value)
                        val = val if val else "NINGUNA"
                        cell.value = limit(val, 10)
                        continue

                    # ---------------- SI / NO ----------------
                    if col_num in [48, 49, 96]:
                        val = str(cell_value or '').strip().upper()
                        cell.value = "SI" if val in ['SI', 'TRUE', '1', 'S'] else "NO"
                        continue

                    # ---------------- CLASE RIESGO ----------------
                    if col_num == 55:
                        val = str(cell_value or '').strip()
                        cell.value = "Sin Riesgo" if val in ['', '5', 'NONE', 'FALSE', '0'] else val
                        continue

                    # ---------------- CENTRO TRABAJO ----------------
                    if col_num == 79:
                        val = clean_text(cell_value)
                        val = val if val else "RIESGO III"
                        cell.value = limit(val, 10)
                        cell.data_type = 's'
                        continue

                    # ---------------- ACTIVIDAD ECONOMICA ----------------
                    if col_num == 80:
                        val = clean_text(cell_value)
                        cell.value = limit(val, 6)
                        cell.data_type = 's'
                        continue

                    # ---------------- DINERO ----------------
                    if col_num in [47, 52, 54, 56, 57, 58, 59, 60, 61, 65, 68, 70, 72, 76, 81, 84, 86, 87, 89, 91, 93, 95]:
                        try:
                            num = float(cell_value or 0)
                            if num != 0:
                                cell.value = int(num)
                                cell.number_format = '#,##0'
                            else:
                                cell.value = "" # <--- Cambiado para que quede en blanco si es 0
                            cell.alignment = Alignment(horizontal='right', vertical='center')
                        except:
                            cell.value = ""
                        continue


                    # ---------------- PORCENTAJES ----------------
                    if col_num in [53, 66, 77, 85, 88, 90, 92, 94]:
                        try:
                            val = float(str(cell_value or 0).replace('%', '').strip())
                            cell.value = val / 100 if val > 1 else val
                            cell.number_format = '0.00%'
                        except:
                            cell.value = ""
                        continue

                    # ---------------- FECHAS ----------------
                    if col_num in [15, 17, 19, 25, 28, 29, 31, 32, 34, 35, 37, 38, 41, 42, 44, 45]:
                        if cell_value:
                            try:
                                # Si ya es date/datetime, lo usamos. Si es string, lo convertimos.
                                if isinstance(cell_value, (date, datetime)):
                                    val_fecha = cell_value
                                else:
                                    val_fecha = fields.Date.from_string(cell_value)

                                cell.value = val_fecha
                                cell.number_format = 'yyyy-mm-dd' # Estándar requerido por PILA
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                            except:
                                cell.value = ""
                        else:
                            cell.value = ""
                        continue

                    # ---------------- NUMÉRICOS GENERALES ----------------
                    if (47 <= col_num <= 78) or (81 <= col_num <= 95) or col_num == 12:
                        try:
                            num = float(cell_value or 0)
                            if num != 0:
                                cell.value = int(num)
                                cell.number_format = '#,##0'
                            else:
                                cell.value = ""
                        except:
                            cell.value = ""
                        continue

                    # ---------------- DEFAULT ----------------

                    # Este código solo se ejecutará si NO entró en ninguno de los IF anteriores
                    val = clean_text(cell_value)
                    cell.value = limit(val, 30)
                    cell.alignment = Alignment(horizontal='left', vertical='center')

                # --- ESTAS DOS LÍNEAS DEBEN IR AQUÍ (Fuera del bucle de columnas) ---
                row_num += 1
            row_counter += 1

        # --- TODO ESTO DEBE IR FUERA DEL BUCLE DE EMPLEADOS ---
        #last_data_row = row_num - 1

        #if sheet.max_row > last_data_row:
            #sheet.delete_rows(last_data_row + 1, sheet.max_row - last_data_row)

        # 1. Guardamos el archivo directamente (SIN borrar filas sobrantes)
        # 1. Guardar en el buffer
        wb.save(output)

        # 2. OBTENER el valor PRIMERO
        out_data = output.getvalue()

        # 3. CERRAR el buffer DESPUÉS
        output.close()

        # 4. Crear el adjunto
        attachment = self.env['ir.attachment'].create({
            'name': f"PILA_DIVITIASSAS_{fields.Date.today()}.xlsx",
            'type': 'binary',
            'datas': base64.b64encode(out_data), # Aquí ya tenemos los datos seguros
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })

        # 3. Retornamos la descarga
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }



    def _format_pila(self, value, length, field_type='text'):
        """ Limpieza de datos y ajuste de ancho fijo """
        val = str(value).strip().upper() if value and not pd.isna(value) else ""

        # Lógica para Novedades (IGE, LMA, IRP) - Posición 152 corregida
        if length == 2 and field_type == 'text':
            if val in ['SI', 'S', 'X', '1', 'TRUE']:
                return 'X '.ljust(2)
            return '  '

        if field_type == 'num':
            # Deja solo números y rellena con ceros a la izquierda
            clean_num = "".join(filter(str.isdigit, val))
            return clean_num.zfill(length)

        # Rellena con espacios a la derecha para texto
        return val.ljust(length)[:length]


    def action_generate_txt_pila(self):
        """ Convierte el Excel cargado en un archivo TXT plano para Aportes en Línea """
        self.ensure_one()
        if not self.plantilla_excel:
            raise UserError(_("Por favor, suba primero el archivo Excel del cliente."))

        # 1. Leer el archivo cargado
        file_data = base64.b64decode(self.plantilla_excel)
        try:
            try:
                df = pd.read_excel(io.BytesIO(file_data))
            except:
                df = pd.read_csv(io.BytesIO(file_data))
        except Exception as e:
            raise UserError(_("No se pudo leer el archivo. Asegúrese de que sea un Excel válido. Error: %s") % str(e))

        lines = []
        for _, row in df.iterrows():
            # --- CONSTRUCCIÓN DEL REGISTRO TIPO 2 (ANCHO FIJO 298) ---
            line = "02" # Registro Tipo 2 (Pos 1-2)

            # Documento (Pos 3-18) - Buscamos el nombre de columna en tu excel viejo
            identificacion = row.get('No. Id') or row.get('Identificación') or ""
            line += self._format_pila(identificacion, 16, 'num')

            # Espacios hasta llegar a la posición 147
            line = line.ljust(147)

            # --- SECCIÓN DE NOVEDADES (POSICIONES CRÍTICAS) ---
            # Pos 148-149: IGE
            line += self._format_pila(row.get('Incapacidad General'), 2)
            # Pos 150-151: LMA
            line += self._format_pila(row.get('Licencia Maternidad Paternidad'), 2)
            # Pos 152-153: IRP (¡Aquí se soluciona el error de Aportes en Línea!)
            line += self._format_pila(row.get('Incapacidad Riesgos Profesionales'), 2)

            # --- FINALIZACIÓN ---
            # Rellenar con espacios hasta completar los 298 caracteres exactos
            line = line.ljust(298)
            lines.append(line)

        # Unir líneas y convertir a Base64
        final_txt = "\r\n".join(lines)
        attachment_name = 'PILA_APORTES_%s.txt' % self.date_to

        attachment = self.env['ir.attachment'].create({
            'name': attachment_name,
            'type': 'binary',
            'datas': base64.b64encode(final_txt.encode('latin-1')), # Latin-1 para formato ANSI
            'mimetype': 'text/plain',
        })

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }
