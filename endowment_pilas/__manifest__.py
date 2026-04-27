{
    'name': 'dotacion_nomina',
    'version': '18.0.2.0.3',
    'summary': 'Integra los datos de las administradoras y conceptos',
    'description': 'Integra los datos de las administradoras y conceptos',
    'category': 'Nomina',
    'author': 'Navegasoft ,Colaborador:Ing.Marilynmillan.',
    'depends': ['base','stock','hr','hr_payroll','hr_contract','hr_holidays','fix_extructuras_de_pago'],
    'data': [
        'security/ir.model.access.csv',
        'views/pila_novelty_type.xml',
        'views/hr_administradoras_views.xml',
        'views/list_administradora.xml',
        'views/hr_contract.xml',
        'views/hr_contract_concept.xml',
        'views/hr_leave_type.xml',
        'views/hr_employee.xml',
        'views/hr_payslip.xml',
        'views/hr_salary_rule.xml',
        'views/work_center.xml',
        'views/pila_tipo_trabajador_views.xml',
        'views/pila_subtipo_trabajador_views.xml',
        'wizard/payrol_excel_wizard_view.xml',



    ],


    'demo': [],
    'license': 'LGPL-3',
    'application': True,
}
# -*- coding: utf.8 -*-
