from models.contact import Contact


class Sender(Contact):
    def __init__(self, name: str, email: str, title: str, company: str, phone: str):
        super().__init__(name, email)
        self.title = title
        self.company = company
        self.phone = phone

    def get_signature(self):
        return '<br><br><small>Kind regards,</small>' \
            f'<br><strong>{self.name}</strong>' \
            f'<br><small>{self.title} <strong>{self.company}</strong></small>' \
            f'<br><small>{self.phone}</small>'
