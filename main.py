from docx import Document
from appscript import app, k
from mactypes import Alias
from pathlib import Path
from docx2pdf import convert

from docx_replace import docx_replace

invoiceNumber = '013'
invoiceMonth = 'February'
invoiceDay = '28'
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

content = f'Hello.<br><br>I hope you\'re well.' \
          f'<br><br>Please find attached the invoice number {invoiceNumber} for Base Pay for {invoiceMonth}, {invoiceYear}.'

signature = '<br><br><small>Kind regards,</small>' \
            '<br><strong>Elias Matheus Melo de Oliveira</strong>' \
            '<br><small>Founder <strong>Olitech</strong></small>' \
            '<br><small>+55 21 99988-7172</small>'

msg = outlook.make(
    new=k.outgoing_message,
    with_properties={
        k.subject: f'Invoice {invoiceNumber} for {invoiceMonth} - {invoiceYear}',
        k.content: content + signature})

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
