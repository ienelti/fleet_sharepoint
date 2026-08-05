# -*- coding: utf-8 -*-
import base64
import logging
import requests
import os
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

    # 1. Restricción: Evita que el mismo vehículo tenga dos documentos con el mismo 'name'
    @api.constrains('name', 'vehicle_id')
    def _check_unique_document_name(self):
        for doc in self:
            if not doc.name or not doc.vehicle_id:
                continue
                
            # Buscamos si existe otro documento en el mismo vehículo con el mismo nombre
            # Usamos '=ilike' para que sea insensible a mayúsculas/minúsculas 
            # (ej. "SOAT" será tratado igual que "soat")
            domain = [
                ('vehicle_id', '=', doc.vehicle_id.id),
                ('name', '=ilike', doc.name),
                ('id', '!=', doc.id) # Excluimos el registro actual que se está guardando
            ]
            
            # Si el contador encuentra al menos 1 coincidencia, bloqueamos el guardado
            if self.env['fleet.sharepoint.document'].search_count(domain) > 0:
                raise exceptions.ValidationError(
                    f"El documento con el nombre '{doc.name}' ya existe para este vehículo.\n"
                    "Por favor, asigne un nombre diferente para evitar sobreescribir el archivo en SharePoint."
                )

    def _get_sharepoint_credentials(self):
        ICPSudo = self.env['ir.config_parameter'].sudo()
        tenant_id = ICPSudo.get_param('fleet_sharepoint.tenant_id')
        client_id = ICPSudo.get_param('fleet_sharepoint.client_id')
        client_secret = ICPSudo.get_param('fleet_sharepoint.client_secret')
        drive_id = ICPSudo.get_param('fleet_sharepoint.drive_id')

        if not all([tenant_id, client_id, client_secret, drive_id]):
            raise exceptions.UserError("Faltan credenciales de Microsoft SharePoint en los Ajustes.")
        return tenant_id, client_id, client_secret, drive_id

    def _get_access_token(self, tenant_id, client_id, client_secret):
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        payload = {
            'client_id': client_id,
            'scope': 'https://graph.microsoft.com/.default',
            'client_secret': client_secret,
            'grant_type': 'client_credentials'
        }
        response = requests.post(url, data=payload, timeout=10)
        if response.status_code != 200:
            raise exceptions.UserError("No se pudo autenticar con Microsoft 365.")
        return response.json().get('access_token')

    def _upload_to_sharepoint(self, file_b64, final_filename, vehicle):
        tenant_id, client_id, client_secret, drive_id = self._get_sharepoint_credentials()
        access_token = self._get_access_token(tenant_id, client_id, client_secret)

        folder_name = vehicle.license_plate or vehicle.name or "Vehiculo_Sin_Identificador"
        
        # Limpiamos el nombre final por seguridad
        clean_filename = final_filename.replace(" ", "_").replace("/", "-")

        # Tu ruta exacta en SharePoint
        base_path = "/08 TI/04 Desarrollos/04 Odoo/Odoo19/Creacion de modulos/fleet_sharepoint/documentos_odoo"
        
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{base_path}/{folder_name}/{clean_filename}:/content"

        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/octet-stream'
        }

        file_data = base64.b64decode(file_b64)
        response = requests.put(url, headers=headers, data=file_data, timeout=30)

        if response.status_code in [200, 201]:
            res_data = response.json()
            item_id = res_data.get('id')
            mime_type = res_data.get('file', {}).get('mimeType', 'application/pdf')
            return item_id, mime_type
        else:
            raise exceptions.UserError(f"Error al subir a SharePoint: {response.status_code} - {response.text}")

    @api.model_create_multi
    def create(self, vals_list):
        for vals in vals_list:
            if vals.get('upload_file'):
                file_data = vals.get('upload_file')
                
                # 2. Lógica de Renombramiento Inteligente
                # Extraemos la extensión del archivo original (ej. .pdf)
                original_filename = vals.get('filename', 'documento.pdf')
                _, ext = os.path.splitext(original_filename)
                if not ext:
                    ext = '.pdf'
                
                # Tomamos el nombre ingresado por el usuario y le concatenamos la extensión
                document_name = vals.get('name', 'Documento')
                final_filename = f"{document_name}{ext}"
                
                vehicle_id = self.env['fleet.vehicle'].browse(vals.get('vehicle_id'))
                
                # Pasamos el final_filename a SharePoint
                sp_id, mime = self._upload_to_sharepoint(file_data, final_filename, vehicle_id)
                
                vals['sp_item_id'] = sp_id
                vals['mimetype'] = mime
                vals['upload_file'] = False
                
        return super(FleetSharepointDocument, self).create(vals_list)

    def action_view_document(self):
        self.ensure_one()
        if not self.sp_item_id:
            raise exceptions.UserError("Este documento no tiene un archivo asociado en SharePoint.")
            
        wizard = self.env['fleet.sharepoint.viewer.wizard'].create({
            'document_id': self.id
        })

        return {
            'name': f'Documento: {self.name}',
            'type': 'ir.actions.act_window',
            'res_model': 'fleet.sharepoint.viewer.wizard',
            'res_id': wizard.id,
            'view_mode': 'form',
            'target': 'new',
        }

    def action_download_document(self):
        """ Descarga física del archivo """
        self.ensure_one()
        return {
            'type': 'ir.actions.act_url',
            # Apuntamos a una nueva ruta de descarga en el controlador
            'url': f'/sharepoint/download/{self.id}',
            'target': 'self', # 'self' hace que inicie la descarga sin abrir pestañas nuevas
        }

    def action_share_document(self):
        """ Genera link anónimo de solo lectura en SharePoint """
        self.ensure_one()
        tenant_id, client_id, client_secret, drive_id = self._get_sharepoint_credentials()
        access_token = self._get_access_token(tenant_id, client_id, client_secret)
        
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{self.sp_item_id}/createLink"
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        # Parámetros: 'view' (Solo lectura), 'anonymous' (Cualquiera con el link)
        payload = {"type": "view", "scope": "anonymous"}
        
        response = requests.post(url, headers=headers, json=payload, timeout=20)
        
        if response.status_code in [200, 201]:
            link_url = response.json().get('link', {}).get('webUrl')
            
            # Abrimos un modal para mostrarle el enlace al usuario
            wizard = self.env['fleet.sharepoint.share.wizard'].create({'share_url': link_url})
            return {
                'name': 'Enlace de Compartición',
                'type': 'ir.actions.act_window',
                'res_model': 'fleet.sharepoint.share.wizard',
                'res_id': wizard.id,
                'view_mode': 'form',
                'target': 'new',
            }
        else:
            raise exceptions.UserError(f"No se pudo generar el enlace. SharePoint respondió: {response.text}")

    def write(self, vals):
        # 1. Guardamos el cambio en Odoo primero
        res = super(FleetSharepointDocument, self).write(vals)
        
        # 2. Si el usuario cambió el campo 'name', le avisamos a SharePoint
        if 'name' in vals:
            for doc in self:
                if doc.sp_item_id:
                    tenant_id, client_id, client_secret, drive_id = self._get_sharepoint_credentials()
                    access_token = self._get_access_token(tenant_id, client_id, client_secret)
                    
                    # Recuperamos la extensión original para no corromper el archivo
                    _, ext = os.path.splitext(doc.filename or '.pdf')
                    if not ext:
                        ext = '.pdf'
                    
                    nuevo_nombre_archivo = f"{doc.name}{ext}".replace(" ", "_").replace("/", "-")
                    
                    # Graph API PATCH para renombrar
                    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{doc.sp_item_id}"
                    headers = {
                        'Authorization': f'Bearer {access_token}',
                        'Content-Type': 'application/json'
                    }
                    payload = {"name": nuevo_nombre_archivo}
                    requests.patch(url, headers=headers, json=payload, timeout=20)
                    
                    # Actualizamos también el nombre técnico interno en Odoo
                    doc.filename = nuevo_nombre_archivo
        return res

    def unlink(self):
        # IMPORTANTE: Eliminamos en SharePoint ANTES de borrar el registro en Odoo
        for doc in self:
            if doc.sp_item_id:
                try:
                    tenant_id, client_id, client_secret, drive_id = self._get_sharepoint_credentials()
                    access_token = self._get_access_token(tenant_id, client_id, client_secret)
                    
                    # Graph API DELETE para eliminar el archivo físicamente
                    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{doc.sp_item_id}"
                    headers = {'Authorization': f'Bearer {access_token}'}
                    requests.delete(url, headers=headers, timeout=20)
                except Exception as e:
                    _logger.error("Error al intentar borrar archivo en SharePoint: %s", str(e))
                    
        # Luego borramos de la base de datos de Odoo
        return super(FleetSharepointDocument, self).unlink()