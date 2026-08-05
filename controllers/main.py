# -*- coding: utf-8 -*-
import requests
import logging
import zipfile
import io
from odoo import http
from odoo.http import request

_logger = logging.getLogger(__name__)

class SharepointController(http.Controller):

    @http.route('/sharepoint/document/<int:doc_id>', type='http', auth='user', website=False)
    def download_sharepoint_document(self, doc_id, **kwargs):
        """ Endpoint que sirve de proxy entre SharePoint y el visor de Odoo """
        # 1. Buscamos el documento registrado en Odoo
        doc = request.env['fleet.sharepoint.document'].sudo().browse(doc_id)
        if not doc.exists() or not doc.sp_item_id:
            return request.not_found()

        # 2. Obtenemos credenciales y token
        try:
            tenant_id, client_id, client_secret, drive_id = doc._get_sharepoint_credentials()
            access_token = doc._get_access_token(tenant_id, client_id, client_secret)
        except Exception as e:
            _logger.error("Error al obtener credenciales/token en Controller: %s", str(e))
            return request.make_response("Error de autenticación con SharePoint", status=500)

        # 3. Solicitamos el contenido binario del archivo a Graph API
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{doc.sp_item_id}/content"
        headers = {'Authorization': f'Bearer {access_token}'}

        response = requests.get(url, headers=headers, timeout=30)

        if response.status_code != 200:
            _logger.error("Error al descargar archivo de SharePoint: %s", response.text)
            return request.make_response("No se pudo obtener el archivo desde SharePoint", status=response.status_code)

        # 4. Retornamos la respuesta HTTP con el PDF para visualización Inline
        content_type = doc.mimetype or 'application/pdf'
        filename = doc.filename or f"{doc.name}.pdf"

        headers_http = [
            ('Content-Type', content_type),
            ('Content-Disposition', f'inline; filename="{filename}"'),
            ('Content-Length', len(response.content))
        ]

        return request.make_response(response.content, headers=headers_http)

    @http.route('/sharepoint/download/<int:doc_id>', type='http', auth='user', website=False)
    def download_attachment_sharepoint(self, doc_id, **kwargs):
        doc = request.env['fleet.sharepoint.document'].sudo().browse(doc_id)
        if not doc.exists() or not doc.sp_item_id:
            return request.not_found()

        tenant_id, client_id, client_secret, drive_id = doc._get_sharepoint_credentials()
        access_token = doc._get_access_token(tenant_id, client_id, client_secret)

        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{doc.sp_item_id}/content"
        headers = {'Authorization': f'Bearer {access_token}'}
        response = requests.get(url, headers=headers, timeout=30)

        content_type = doc.mimetype or 'application/octet-stream'
        filename = doc.filename or f"{doc.name}.pdf"

        headers_http = [
            ('Content-Type', content_type),
            # 'attachment' fuerza la descarga en la computadora del usuario
            ('Content-Disposition', f'attachment; filename="{filename}"'),
            ('Content-Length', str(len(response.content)))
        ]
        return request.make_response(response.content, headers=headers_http)

    @http.route('/sharepoint/download_zip', type='http', auth='user', website=False)
    def download_zip_sharepoint(self, ids, **kwargs):
        """ Recibe múltiples IDs, los descarga de SharePoint y devuelve un ZIP """
        if not ids:
            return request.not_found()

        # Convertimos la lista de IDs (ej: "1,2,3") en una lista de enteros
        doc_ids = [int(i) for i in ids.split(',') if i]
        docs = request.env['fleet.sharepoint.document'].sudo().browse(doc_ids)

        # Creamos un archivo ZIP en memoria (RAM)
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for doc in docs:
                if doc.sp_item_id:
                    tenant_id, client_id, client_secret, drive_id = doc._get_sharepoint_credentials()
                    access_token = doc._get_access_token(tenant_id, client_id, client_secret)
                    
                    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{doc.sp_item_id}/content"
                    headers = {'Authorization': f'Bearer {access_token}'}
                    response = requests.get(url, headers=headers, timeout=30)
                    
                    if response.status_code == 200:
                        # Guardamos el archivo dentro del ZIP con su nombre correspondiente
                        filename = doc.filename or f"{doc.name}.pdf"
                        zip_file.writestr(filename, response.content)

        # Movemos el cursor de lectura al inicio del archivo en memoria
        zip_buffer.seek(0)

        # Retornamos el ZIP al navegador
        headers_http = [
            ('Content-Type', 'application/zip'),
            ('Content-Disposition', 'attachment; filename="Documentos_Flota.zip"'),
            ('Content-Length', str(len(zip_buffer.getvalue())))
        ]
        
        return request.make_response(zip_buffer.getvalue(), headers=headers_http)