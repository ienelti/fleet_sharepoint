# -*- coding: utf-8 -*-
from odoo import models, fields

class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'

    # Usamos config_parameter para que Odoo lo guarde en la tabla ir.config_parameter
    sharepoint_tenant_id = fields.Char(
        string="Tenant ID",
        config_parameter='fleet_sharepoint.tenant_id',
        help="ID del Tenant de Microsoft 365"
    )
    sharepoint_client_id = fields.Char(
        string="Client ID",
        config_parameter='fleet_sharepoint.client_id',
        help="ID de la aplicación (Client ID) registrada en Azure/Entra ID"
    )
    sharepoint_client_secret = fields.Char(
        string="Client Secret",
        config_parameter='fleet_sharepoint.client_secret',
        help="Secreto de la aplicación de Azure"
    )
    sharepoint_drive_id = fields.Char(
        string="Drive ID / Site ID",
        config_parameter='fleet_sharepoint.drive_id',
        help="ID del Drive de SharePoint donde se creará la estructura de carpetas"
    )