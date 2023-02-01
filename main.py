from docx import Document
from appscript import app, k
from mactypes import Alias
from pathlib import Path
from docx2pdf import convert

from docx_replace import docx_replace

invoiceNumber = '012'
invoiceMonth = 'January'
invoiceDay = '31'
invoiceYear = '2023'

invoiceDate = f'{invoiceMonth} {invoiceDay}, {invoiceYear}'

# usage
doc = Document('./template/template.docx')
docx_replace(doc, dict(InvoiceNumber=invoiceNumber, InvoiceDate=invoiceDate, InvoiceMonth=invoiceMonth))
doc.save('result/invoice.docx')

# convert to PDF
convert("result/invoice.docx", "result/invoice.pdf")

outlook = app('Microsoft Outlook')

p = Path('result/invoice.pdf')
p = Alias(str(p))

msg = outlook.make(
    new=k.outgoing_message,
    with_properties={
        k.subject: f'Invoice for {invoiceMonth} - {invoiceYear}',
        k.plain_text_content: 'Hello. \nPlease find attached invoice'})

msg.make(
    new=k.recipient,
    with_properties={
        k.email_address: {
            k.name: 'Payments Rainforest',
            k.address: 'payment@rainforest.tech'}})

msg.make(
    new=k.attachment,
    with_properties={
        k.file: p}
)

msg.open()
msg.activate()

