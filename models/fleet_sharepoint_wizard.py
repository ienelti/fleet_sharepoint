# -*- coding: utf-8 -*-
from odoo import models, fields, api

class FleetSharepointViewerWizard(models.TransientModel):
    _name = 'fleet.sharepoint.viewer.wizard'
    _description = 'Visor Modal de Documentos SharePoint'

    # Relacionamos el wizard con el documento que queremos ver
    document_id = fields.Many2one('fleet.sharepoint.document', string="Documento")
    
    # Campo HTML donde inyectaremos el visor. 
    # IMPORTANTE: sanitize=False permite que Odoo renderice la etiqueta <iframe> sin borrarla.
    viewer_html = fields.Html(compute='_compute_viewer_html', sanitize=False)

    @api.depends('document_id')
    def _compute_viewer_html(self):
        for wizard in self:
            if wizard.document_id:
                # Armamos la ruta web hacia el controlador que ya existe
                url = f"/sharepoint/document/{wizard.document_id.id}"
                
                # Creamos un iframe que ocupará todo el ancho y tendrá 600px de alto
                wizard.viewer_html = f'''
                    <iframe src="{url}" 
                            width="100%" 
                            height="600px" 
                            style="border: none; border-radius: 5px;">
                    </iframe>
                '''
            else:
                wizard.viewer_html = '<p>No hay documento disponible.</p>'