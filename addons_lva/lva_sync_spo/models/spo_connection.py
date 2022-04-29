from odoo import api, fields, models
from odoo.addons.lva_sync_spo.odoo_sync.config_and_connect_odoo_sync import *
from odoo.addons.lva_sync_spo.odoo_sync.odoo_connect import *
from odoo.addons.lva_sync_spo.odoo_sync.sftp_filemove import *
from odoo.addons.lva_sync_spo.odoo_sync.spo_connect import *


class SPOConnection(models.Model):
    _name = 'spo.connect'

    user_of_spo = fields.Char('Username of SPO')
    password_of_spo = fields.Char('Password of SPO')
    website_spo = fields.Char('Domain of SPO')
    library = fields.Char('SPO Library')

    def get_spo_connection(self, company):
        """
        Returns the SPO Opportunities Connection object for a company.

        Returns: list first value true/false, second value message, third value the connection object
        """
        __spo_conn = self.init_spo_connection()
        if company in __spo_conn:
            return __spo_conn[company]['conn']
        else:
            return False, f'Mistake! There is no SPO Opportunities connection for the company with ID {company}!'

    def read_spo_config(self):
        """
        Reads the configurations for the SPO connections for the Opportunities libraries.

        The SPO configurations are output as self.__spo_crm_conf[<ID>].
         The ID corresponds to the company ID in Oddo. The ID=0 specifies the default values
        """
        # Determine default values for SPO configurations
        conf_error = False
        __config = {'user': self.user_of_spo, 'pass': self.password_of_spo, 'website': self.website_spo, 'library': self.library}

        spo_default_conf = __config
        # Determine if CRM/Sales Sync is turned off by default

        # Determine further default values only if Sync=True
        if 'user' in spo_default_conf:
            default_user = spo_default_conf['user']
        else:
            conf_error = True
            conf_error_msg = (
                "No user specified in the [SPO] section!"
            )
        if 'pass' in spo_default_conf:
            default_pass = spo_default_conf['pass']
        else:
            conf_error = True
            conf_error_msg = (
                "No password entered in the [SPO] section!"
            )
        if 'website' in spo_default_conf:
            default_website = spo_default_conf['website']
        else:
            conf_error = True
            conf_error_msg = (
                "No CRM website specified in the [SPO] section!"
            )
        if 'library' in spo_default_conf:
            default_library = spo_default_conf['library']
        else:
            conf_error = True
            conf_error_msg = (
                "No CRM library specified in the [SPO] section!"
            )

        # Determine configurations for the individual companies
        __spo_conf = {}
        for rec in self.env.user.company_id:
            firma = rec.id
            __spo_conf[firma] = {}

            # Determine if separate configuration
            # available for the corresponding FA
            # Determine if default values are complete
            if not conf_error:
                __spo_conf[firma]['user'] = default_user
                __spo_conf[firma]['pass'] = default_pass
                __spo_conf[firma]['website'] = default_website
                __spo_conf[firma]['library'] = default_library
        return __spo_conf

    def get_spo_config(self, company):
        """
        Returns the configuration data for the library
         of machine attachments for a company.

        Rückgabe CRM/Sales-Conf für FA
        """
        __spo_conf = self.read_spo_config()
        return __spo_conf[company]

    def init_spo_connection(self):
        """
        Initializes the accessors for the
        Library/s for the projects

        A dictionary with the following structure is created:

        self.__spo_crm_conn[<ID>]['status']      True/False indicates whether a
                                                 Connection object to successfully
                                                 was set up.
        self.__spo_crm_conn[<ID>]['message']    Text. Gives the error/ok message for
                                                 the DB connection object
        self.__spo_crm_conn[<ID>]['connection'] The actual DB connection object
        """

        # Get Config for Libraries
        __spo_conf = self.read_spo_config()
        # Process the configurations of the individual companies
        # and create Connection objects
        __spo_conn = {}
        for conf in __spo_conf:
            __spo_conn[conf] = {}
            __spo_conn[conf]['conn'] = SPOLibrary(
                __spo_conf[conf]['website'], __spo_conf[conf]['library'], __spo_conf[conf]['user'],
                __spo_conf[conf]['pass'])
            # Only create a connection object if required!
            result = __spo_conn[conf]['conn'].test_connection()
            if result[0]:
                __spo_conn[conf]['status'] = True
                __spo_conn[conf]['message'] = 'Connection can be established!'
            else:
                __spo_conn[conf]['status'] = False
                __spo_conn[conf]['message'] = 'Mistake! Connection cannot be established with the configuration!'
        return __spo_conn