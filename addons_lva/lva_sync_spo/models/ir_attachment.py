from odoo import api, fields, models, _
import os
import shutil
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from odoo.exceptions import UserError


class IrAttachment(models.Model):
    _inherit = 'ir.attachment'

    def _compute_mimetype(self, values):
        if values.get("url"):
            return "application/link"
        return super()._compute_mimetype(values)

    @api.model
    def create(self, vals_list):
        res = super(IrAttachment, self).create(vals_list)
        if 'res_model' in vals_list and 'res_id' in vals_list and self._context.get('spo_sync'):
            self.sync_to(vals_list=vals_list, record=res)
        return res

    def get_record_folder_name(self, model, res_id):
        record = self.env[model].search([('id', '=', res_id)])
        return record

    def sync_file_to_spo(self, record_id, res_model, att_id, conf_conn):
        if conf_conn:
            spo_conf = conf_conn.get_spo_config(self.env.user.company_id.id)
            spo_conn = conf_conn.get_spo_connection(self.env.user.company_id.id)
            folder_name = f"[{record_id.id}]_{record_id.name}".replace('/', '_'). \
                replace('#', '_').replace(':', '_').replace('\\', '_').strip().replace(' ', '_')
            folder_exist = False
            folder_link = None
            folder_id = None
            result = spo_conn.get_folder_by_path(f"/{folder_name}")
            if not result[0]:
                result = spo_conn.create_folder(f"/{folder_name}")
                result = spo_conn.get_folder_by_path(f"/{folder_name}")
                if result[0]:
                    folder_id = result[1]._properties['ID']
                    folder_exist = True
            else:
                folder_id = result[1]._properties['ID']
                folder_exist = True
            if folder_exist:
                result = spo_conn.get_item_by_id(folder_id, ["FileLeafRef", "FileRef", "EncodedAbsUrl"])
                if result[0]:
                    folder_link = result[1]._properties['EncodedAbsUrl']
                else:
                    folder_link = f"{spo_conf['website']}/{spo_conf['library']}/{folder_name}"
            else:
                print(
                    f"Mistake! No folder/docset could be found or created for record {record_id.id}")
            if folder_exist:
                attachment = self.env['ir.attachment'].search([('id', '=', att_id)])
                if attachment:
                    if len(attachment) > 0:
                        # -> transfer attachments
                        source_path, file_store_path = self.copy_file_path(attachment=attachment)
                        target_path = f'{folder_name}/{attachment.name}'
                        spo_conn.upload_file_to_library(source_path, target_path)
                        file_link = f'{folder_link}/{attachment.name}'
                        return file_link, attachment.name, source_path, file_store_path
                    else:
                        print(
                            f"-> There are no attachments for the transfer to the SPO for the record {record_id.id}.")
                else:
                    print(
                        f"Mistake! An error occurred while querying the attachments for the record {record_id.id}!\n{res_model}")
            else:
                print(
                    f"Mistake! Cannot copy the attachments to the record {record_id.id} into the SPO because there is no folder/docset for it!")
        else:
            pass

    def get_attachment_folder_link(self, record_id, model_name):
        result = self.env['ir.attachment'].search([('res_model', '=', model_name), ('res_id', '=', record_id.id)])
        if result:
            return True, result

    def copy_file_path(self, attachment):
        res = attachment._full_path(attachment.store_fname)
        res_revert = res[::-1]
        index = res_revert.find('/')
        res_revert = res_revert[index:]
        destination = str(res_revert[::-1]) + str(attachment.name)
        try:
            shutil.copyfile(res, destination)

        except shutil.SameFileError:
            print("Source and destination represents the same file.")

        # If destination is a directory.
        except IsADirectoryError:
            print("Destination is a directory.")

        # If there is any permission issue
        except PermissionError:
            print("Permission denied.")
        # For other errors
        except:
            print("Error occurred while copying file.")
        return destination, res

    def delete_file_store(self, abspath):
        try:
            os.remove(abspath)
            return True
        except:
            print('not work')
        return True

    def sync_to(self, vals_list, record):
        conf_conn = self.env['spo.connect'].search([], limit=1)
        if conf_conn:
            record_id = self.get_record_folder_name(model=vals_list['res_model'], res_id=int(vals_list['res_id']))
            file_link, name_file, source_path, file_store_path = self.sync_file_to_spo(record_id=record_id, res_model=vals_list['res_model'], att_id=record.id,
                                                                                       conf_conn=conf_conn)
            if file_link and name_file and source_path and file_store_path:
                record.url = file_link
                record.name = name_file
                record.type = 'url'
                record.datas = False
                record.mimetype = 'application/link'
                self.delete_file_store(file_store_path)
                self.delete_file_store(source_path)
