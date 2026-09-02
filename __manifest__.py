{
    "name": "DCL Debtors Age Analysis",
    "summary": "Convert the monthly debtors ledger into a property-by-property aged receivables workbook",
    "version": "19.0.1.0.0",
    "category": "Accounting/Reporting",
    "author": "Dunhill Consulting Limited",
    "license": "LGPL-3",
    "depends": ["base", "web", "account"],
    "external_dependencies": {
        "python": ["openpyxl", "xlsxwriter"],
    },
    "data": [
        "security/ir.model.access.csv",
        "views/debtors_age_wizard_views.xml",
        "views/menu_views.xml",
    ],
    "installable": True,
    "application": True,
}
