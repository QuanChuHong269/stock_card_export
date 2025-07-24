{
    'name': 'Stock Card Export',
    'version': '1.0',
    'description': 'Xuất Thẻ Kho',
    'depends': ['stock'],
    'category': 'Inventory',
    'data': [
        'security/ir.model.access.csv',
        'wizard/wizard_stock_card_export_view.xml',
    ],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}