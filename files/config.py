import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RESOURCES_DIR = os.path.join(BASE_DIR, 'resources')

BALANCE_FILE = os.path.join(RESOURCES_DIR, 'Updated Balance (1).xlsx')
INVOICE_TEMPLATE = os.path.join(RESOURCES_DIR, 'Invoice Templates copy.xlsx')
CREDENTIALS_PATH = os.path.join(RESOURCES_DIR, 'credentials.json')
TOKEN_PATH = os.path.join(RESOURCES_DIR, 'token.json')
VAT_FORMS_DIR = os.path.join(RESOURCES_DIR, 'Vat Forms')
