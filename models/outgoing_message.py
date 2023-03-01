from models.invoice import Invoice
from models.sender import Sender


class OutgoingMessage:
    """
    Represents an outgoing message.
    :param invoice: Invoice object.
    :param sender: Sender object.
    """

    def __init__(self, invoice: Invoice, sender: Sender):
        self.invoice = invoice
        self.sender = sender

    def get_subject(self):
        """Returns the outgoing message subject."""
        return f'Invoice {self.invoice.number} for {self.invoice.month} - {self.invoice.year}'

    def get_content(self):
        """Returns the outgoing message content."""
        return f'Hello.<br><br>I hope you\'re well.' \
            f'<br><br>Please find attached the invoice number {self.invoice.number} ' \
            f' for Base Pay for {self.invoice.month}, {self.invoice.year}.'

    # Teste
    def get_signature(self):
        """Return signature of sender."""
        return self.sender.get_signature()
