# -*- coding: utf-8 -*-
from odoo import models, fields, api

class FleetVehicle(models.Model):
    _inherit = 'fleet.vehicle'

    # Campo One2many que enlaza con los documentos de SharePoint
    sharepoint_document_ids = fields.One2many(
        'fleet.sharepoint.document', # El modelo al que apuntamos
        'vehicle_id',                # El campo en el otro modelo que nos referencia
        string="Documentos en SharePoint"
    )

    sharepoint_doc_count = fields.Integer(
        string="Cantidad Documentos", 
        compute='_compute_sharepoint_doc_count'
    )

    @api.depends('sharepoint_document_ids')
    def _compute_sharepoint_doc_count(self):
        for vehicle in self:
            vehicle.sharepoint_doc_count = len(vehicle.sharepoint_document_ids)