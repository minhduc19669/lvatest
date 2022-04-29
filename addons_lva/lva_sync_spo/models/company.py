from odoo.addons.lva_sync_spo.odoo_sync.config_and_connect_odoo_sync import *
from odoo.addons.lva_sync_spo.odoo_sync.odoo_connect import *
from odoo.addons.lva_sync_spo.odoo_sync.sftp_filemove import *
from odoo.addons.lva_sync_spo.odoo_sync.spo_connect import *
from odoo import api, fields, models


class Company(models.Model):
    _inherit = 'res.company'

    user_of_spo = fields.Char('Username of SPO')
    password_of_spo = fields.Char('Password of SPO')
    website_spo = fields.Char('Domain of SPO')
    library = fields.Char('SPO Library')

    def check_connection_spo(self):
        spo_conn = self.env['spo.connect'].search([])
        if spo_conn:
            for spo in spo_conn:
                spo.unlink()
        for rec in self:
            test = self.env['spo.connect'].create({
                'user_of_spo': rec.user_of_spo,
                'password_of_spo': rec.password_of_spo,
                'website_spo': rec.website_spo,
                'library': rec.library
            })
        SPO_LIBRARY = SPOLibrary(website=self.website_spo, library=self.library, user=self.user_of_spo, passwd=self.password_of_spo)
        message = SPO_LIBRARY.test_connection()
        view = self.env.ref('sh_message.sh_message_wizard')
        context = dict(self._context or {})
        context['message'] = message
        return {
            'name': 'Notification!',
            'type': 'ir.actions.act_window',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'sh.message.wizard',
            'views': [(view.id, 'form')],
            'view_id': view.id,
            'target': 'new',
            'context': context
        }
