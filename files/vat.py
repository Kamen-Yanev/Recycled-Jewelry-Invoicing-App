import io
from datetime import date
import pdfplumber
import fitz  # pymupdf


def populate_vat(vat_template, description, quantity_mass, pn, grand_total):
    """
    Fills in the VAT form template with invoice data.
    Returns a BytesIO buffer of the filled PDF — never saved to disk.
    """
    doc = fitz.open(vat_template)

    page1 = doc[0]
    with pdfplumber.open(vat_template) as pdf:
        plumb_page1 = pdf.pages[0]
        plumb_page2 = pdf.pages[1]

        coordinates = []
        words1 = plumb_page1.extract_words()
        words2 = plumb_page2.extract_words()

        fields = ['Description', 'Quantity', 'Selling', 'Reference']
        for field in fields:
            for w in words1:
                if field.lower() in w['text'].lower():
                    coordinates.append(w)
        del coordinates[0]

        for w in words2:
            if 'date' in w['text'].lower():
                coordinates.append(w)
                break

    page1_inserts = {
        'Description': (83.11 + 55,     coordinates[0]['bottom']),
        'Quantity':    (83.11 + 55,     coordinates[1]['bottom']),
        'Selling':     (67.824 + 193.5, coordinates[2]['bottom']),
        'Reference':   (111.40 + 150,   coordinates[3]['bottom'])
    }
    page2_inserts = {
        'Date': (96 + 85, 233.5)
    }

    page1.insert_text(
        (page1_inserts['Description'][0], page1_inserts['Description'][1]),
        description, fontsize=10
    )
    page1.insert_text(
        (page1_inserts['Quantity'][0], page1_inserts['Quantity'][1]),
        quantity_mass, fontsize=10
    )
    page1.insert_text(
        (page1_inserts['Selling'][0], page1_inserts['Selling'][1]),
        f'R{grand_total:,.2f}', fontsize=10
    )
    page1.insert_text(
        (page1_inserts['Reference'][0], page1_inserts['Reference'][1]),
        f'PN{pn}', fontsize=10
    )

    page2 = doc[1]
    page2.insert_text(
        (page2_inserts['Date'][0], page2_inserts['Date'][1]),
        str(date.today()), fontsize=10
    )

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer
