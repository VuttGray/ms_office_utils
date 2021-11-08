from logging import getLogger
from pytz import timezone

logger = getLogger('logger')


class EmailConfig:
    def __init__(self, **kwargs):
        self.time_zone = kwargs.pop('time_zone')
        self.vip_senders = kwargs.pop('vip_senders', [])
        self.spam_addresses = kwargs.pop('spam_addresses', [])


conf: EmailConfig


def configure_email(**kwargs):
    global conf
    conf = EmailConfig(**kwargs)
    return conf


class Email:
    @property
    def entry_id(self):
        return self.__mail_item.EntryID

    @property
    def id(self):
        return f'EML#{self.__mail_item.EntryID[-9:]}'

    @property
    def to(self):
        if self.__mail_item.MessageClass in ['IPM.Note',
                                             'IPM.Note.Microsoft.Missed.Voice',
                                             'IPM.Note.Rules.OofTemplate.Microsoft',
                                             'IPM.Note.Microsoft.Missed']:
            return self.__mail_item.To
        else:
            return ""

    @property
    def sender_email_type(self):
        return self.__mail_item.SenderEmailType

    @property
    def sender_email_address(self):
        if self.sender_email_type == "SMTP":
            return self.__mail_item.SenderEmailAddress.lower()
        elif self.sender_email_type == "EX":
            return self.__mail_item.SenderEmailAddress
        else:
            return ""

    @property
    def sender_name(self):
        if self.__mail_item.MessageClass == 'IPM.Note':
            return self.__mail_item.Sender.Name
        else:
            return ""

    @property
    def subject(self):
        return self.__mail_item.Subject.strip()

    @property
    def body(self):
        return self.__mail_item.Body

    @property
    def received_time(self):
        try:
            received_time = self.__mail_item.ReceivedTime.replace(tzinfo=None)
        except ValueError as ex:
            if ex.args[0] == 'microsecond must be in 0..999999':
                received_time = self.__mail_item.SentOn.replace(tzinfo=None)
            else:
                received_time = None
                pass
        return timezone(conf.time_zone).localize(received_time)

    @property
    def importance(self):
        return ["Low", "Normal", "High"][self.__mail_item.Importance]

    @property
    def categories(self):
        return self.__mail_item.Categories.split(',')

    @property
    def is_unread(self):
        return self.__mail_item.UnRead

    @property
    def is_vip(self):
        if self.sender_name in conf.vip_senders or self.sender_email_address in conf.vip_senders:
            return True
        return False

    @property
    def is_spam(self):
        return self.sender_email_address in conf.spam_addresses

    def __init__(self, mail_item, folder_name):
        if mail_item is None:
            raise AttributeError("Mail item is none")

        self.__mail_item = mail_item
        self.__folder = folder_name
        self.__is_deleted = False

    def __repr__(self):
        return f"[{self.received_time.strftime('%d-%b-%Y %H:%M')} {self.sender_name} {self.subject} {self.id}]"

    def read(self):
        self.__mail_item.UnRead = False

    def check_move(self, condition, folder, mark_read=True):
        if condition:
            self.move(folder, mark_read)
            return True
        return False

    def move(self, folder, mark_read=True):
        self.__mail_item.UnRead = not mark_read
        self.__folder = folder.Name
        logger.debug(f'Move email {self} to {folder.Name}')
        self.__mail_item.Move(folder)

    def check_delete(self, condition):
        if condition:
            self.delete()
            return True
        return False

    def delete(self):
        logger.debug(f'Delete email {self}')
        self.__mail_item.Delete()
        self.__is_deleted = True

    def forward(self, to):
        mail_item = self.__mail_item.Forward()
        mail_item.To = to
        mail_item.Send()
