# -*- coding: utf-8 -*-
from odoo import models, fields

class FleetVehicle(models.Model):
    _inherit = 'fleet.vehicle'

    # Campo One2many que enlaza con los documentos de SharePoint
    sharepoint_document_ids = fields.One2many(
        'fleet.sharepoint.document', # El modelo al que apuntamos
        'vehicle_id',                # El campo en el otro modelo que nos referencia
        string="Documentos en SharePoint"
    )