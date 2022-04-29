# -*- coding: utf-8 -*-
# This file is part of Odoo. The COPYRIGHT file at the top level of
# this module contains the full copyright notices and license terms.
{
    'name': 'LVA CONNECT SPO',
    'version': '14.0.0.0.0',
    'author': 'Hucke Media GmbH & Co. KG/IFE GmbH',
    'category': 'Custom',
    'website': 'https://www.hucke-media.de/',
    'licence': 'AGPL-3',
    'summary': 'Customisations for LVA',
    'depends': [
        'sh_message',
        'web',
        'mail'
    ],
    'data': [
        'views/company.xml',
        'security/ir.model.access.csv',
        'views/assets.xml'
    ],
    "qweb": ["static/xml/url.xml"],
    'installable': True,
    'application': True,
    'auto_install': False,

}