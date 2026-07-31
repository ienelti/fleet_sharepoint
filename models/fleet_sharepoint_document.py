# -*- coding: utf-8 -*-
import base64
import logging
import requests
from odoo import models, fields, api, exceptions

_logger = logging.getLogger(__name__)

class FleetSharepointDocument(models.Model):
    _name = 'fleet.sharepoint.document'
    _description = 'Documento de Flota en SharePoint'

    vehicle_id = fields.Many2one('fleet.vehicle', string="Vehículo", required=True, ondelete='cascade')
    name = fields.Char(string="Nombre del Documento", required=True)
    upload_file = fields.Binary(string="Subir Archivo", store=False)
    filename = fields.Char(string="Nombre del Archivo")
    
    sp_item_id = fields.Char(string="ID de SharePoint", readonly=True)
    mimetype = fields.Char(string="Tipo Mime", readonly=True)

    def _get_sharepoint_credentials(self):
        """ Obtiene las credenciales configuradas en los Ajustes """
        ICPSudo = self.env['ir.config_parameter'].sudo()
        tenant_id = ICPSudo.get_param('fleet_sharepoint.tenant_id')
        client_id = ICPSudo.get_param('fleet_sharepoint.client_id')
        client_secret = ICPSudo.get_param('fleet_sharepoint.client_secret')
        drive_id = ICPSudo.get_param('fleet_sharepoint.drive_id')

        if not all([tenant_id, client_id, client_secret, drive_id]):
            raise exceptions.UserError(
                "Faltan credenciales de Microsoft SharePoint. "
                "Por favor, configúrelas en los Ajustes de la aplicación de Flotas."
            )
        return tenant_id, client_id, client_secret, drive_id

    def _get_access_token(self, tenant_id, client_id, client_secret):
        """ Obtiene el Token OAuth2 mediante Client Credentials Flow """
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        payload = {
            'client_id': client_id,
            'scope': 'https://graph.microsoft.com/.default',
            'client_secret': client_secret,
            'grant_type': 'client_credentials'
        }
        response = requests.post(url, data=payload, timeout=10)
        if response.status_code != 200:
            _logger.error("Error al autenticar con Azure: %s", response.text)
            raise exceptions.UserError("No se pudo autenticar con Microsoft 365. Verifique sus credenciales.")
        
        return response.json().get('access_token')

    def _upload_to_sharepoint(self, file_b64, filename, vehicle):
        """ Sube el archivo binario a SharePoint usando Microsoft Graph API """
        tenant_id, client_id, client_secret, drive_id = self._get_sharepoint_credentials()
        access_token = self._get_access_token(tenant_id, client_id, client_secret)

        # Usamos la placa del vehículo como carpeta, o el nombre si no tiene placa
        folder_name = vehicle.license_plate or vehicle.name or "Vehiculo_Sin_Identificador"
        # Limpiamos caracteres extraños en el nombre del archivo
        clean_filename = filename.replace(" ", "_")

        # Endpoint de Graph API para subir archivos simples (< 4MB)
        # Ruta en SharePoint: /Flota/{folder_name}/{clean_filename}
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/Flota/{folder_name}/{clean_filename}:/content"

        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/octet-stream'
        }

        # Decodificamos el base64 de Odoo a binario puro
        file_data = base64.b64decode(file_b64)

        response = requests.put(url, headers=headers, data=file_data, timeout=30)

        if response.status_code in [200, 201]:
            res_data = response.json()
            item_id = res_data.get('id')
            mime_type = res_data.get('file', {}).get('mimeType', 'application/pdf')
            return item_id, mime_type
        else:
            _logger.error("Error al subir archivo a SharePoint: %s", response.text)
            raise exceptions.UserError(f"Error al subir el archivo a SharePoint: {response.status_code} - {response.text}")

    @api.model_create_multi
    def create(self, vals_list):
        for vals in vals_list:
            if vals.get('upload_file'):
                file_data = vals.get('upload_file')
                file_name = vals.get('filename', 'documento.pdf')
                vehicle_id = self.env['fleet.vehicle'].browse(vals.get('vehicle_id'))
                
                # Invocamos la subida real a Microsoft Graph
                sp_id, mime = self._upload_to_sharepoint(file_data, file_name, vehicle_id)
                
                vals['sp_item_id'] = sp_id
                vals['mimetype'] = mime
                vals['upload_file'] = False  # Limpiamos para no ocupar espacio en PostgreSQL
                
        return super(FleetSharepointDocument, self).create(vals_list)

    def action_view_document(self):
        self.ensure_one()
        if not self.sp_item_id:
            raise exceptions.UserError("Este documento no tiene un archivo asociado en SharePoint.")
            
        return {
            'type': 'ir.actions.act_url',
            'url': f'/sharepoint/document/{self.id}',
            'target': 'new',
        }