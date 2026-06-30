# Recycled-Jewelry-Invoicing-App
This is a desktop application built specifically for Recycled Jewellery, a South Africa based jewellery and fine metal trading business. The goal is to automate the invoicing workflow and create an invoice database which can later be used for analytic purposes.

Two xlsx files (the CPM document, provided by the company which melts the jewelry) and the invoice template are required in order to use the app.

As of 18.06.2026 the main functionality of the app, namely the automated writing of invoices works.

Next goals are adding a function which sends the invoice automatically to the client's WhatsApp upon generation and creating an SQL database.

As of 24.06.2026 the program also updates the company's Google Drive financial database via the Google Drive and Google Sheets API. Other changes are that it now outputs two files, one PDF (for the database) and one xlsx (will be used for a later project). The program is being actively used and has effectively shortened the time required for invoicing. The sending a WhatsApp message idea has been scrapped for now. 

As of 30.06 the program has been refactored - it has been made more modular and safety measures such as gitignore have been added to protect sensitive information. Furthermore, the function of automatically filling out Vat-264 forms and uploading them to the Google Drive whenever an invoice is created is in the making.
