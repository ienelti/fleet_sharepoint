{
    'name': 'Documentos de Flota en SharePoint',
    'version': '1.0',
    'category': 'Human Resources/Fleet',
    'summary': 'Integra los documentos de vehículos directamente en Microsoft SharePoint',
    'description': """
        Almacena documentos de flota en SharePoint usando Microsoft Graph API.
    """,
    'author': 'IENEL',
    'depends': ['base', 'fleet'],
    'data': [
        'security/ir.model.access.csv',
        'views/res_config_settings_views.xml',
        'views/fleet_vehicle_views.xml',
    ],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}