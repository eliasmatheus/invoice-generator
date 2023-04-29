from docx import Document
from appscript import app, k
from mactypes import Alias
from pathlib import Path
from docx2pdf import convert

from functions.docx_replace import docx_replace
from models.contact import Contact
from models.invoice import Invoice
from models.outgoing_message import OutgoingMessage
from models.sender import Sender


# Modifique apenas abaixo dessa linha
invoice = Invoice('015', 2023, 'April', 29)
sender = Sender(
    'Elias Matheus Melo de Oliveira',
    'eliasmatheus@hotmail.com',
    'Founder',
    'OLITECH LTDA',
    '+55 21 99988-7172')
recipient = Contact('Payments Rainforest', 'payment@rainforest.tech')

# Modifique apenas acima dessa linha


results_path = 'result/invoice'

# usage
doc = Document('./template/template.docx')
docx_replace(doc, dict(
    InvoiceNumber=invoice.number,
    InvoiceDate=invoice.date,
    InvoiceMonth=invoice.month))
doc.save(f'{results_path}-{invoice.number}-{invoice.year}.docx')

# convert to PDF
convert(
    f"{results_path}-{invoice.number}-{invoice.year}.docx",
    f"{results_path}-{invoice.number}-{invoice.year}.pdf")

outlook = app('Microsoft Outlook')

p = Path(f'{results_path}-{invoice.number}-{invoice.year}.pdf')
p = Alias(str(p))


outgoing_message = OutgoingMessage(invoice, sender)
content = outgoing_message.get_content()
signature = outgoing_message.get_signature()
subject = outgoing_message.get_subject()


msg = outlook.make(
    new=k.outgoing_message,
    with_properties={
        k.subject: subject,
        k.content: content + signature})

msg.make(
    new=k.recipient,
    with_properties={
        k.email_address: {
            k.name: recipient.name,
            k.address: recipient.email}})

msg.make(
    new=k.attachment,
    with_properties={
        k.file: p}
)

msg.open()
msg.activate()
