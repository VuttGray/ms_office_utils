from logging import getLogger

from win32com.client.gencache import EnsureDispatch

from ms_office_utils.core import clear_temp_directory

logger = getLogger('logger')


def __create_email(recipients=None, cc_recipients='', subject='', html_body=''):
    outlook = EnsureDispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = recipients
    mail.CC = cc_recipients
    mail.Subject = subject
    mail.HTMLBody = html_body
    return mail


def open_email(recipients=None, cc_recipients='', subject='', html_body=''):
    mail = __create_email(recipients, cc_recipients, subject, html_body)
    mail.Display(True)


def send_email(recipients=None, cc_recipients='', subject='', html_body=''):
    mail = __create_email(recipients, cc_recipients, subject, html_body)
    mail.Send()


class OutlookWrapper:
    @property
    def inbox(self):
        return self.__inbox

    @property
    def sent_items(self):
        return self.__sent_items

    def __init__(self):
        try:
            self.__outlook = EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
            logger.info('Connected to Outlook')
        except AttributeError:
            logger.info('Connection to Outlook failed')
            clear_temp_directory('gen_py', win_dir='TEMP')
            self.__outlook = EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
            logger.info('Connected to Outlook')

        # Default boxes: 3 Deleted Items, 4 Outbox, 5 Sent Items, 6 Inbox,
        # 9 Calendar, 10 Contacts, 11 Journal, 12 Notes, 13 Tasks, 14 Drafts
        self.__inbox = self.__outlook.GetDefaultFolder(6)
        self.__sent_items = self.__outlook.GetDefaultFolder(5)

    @staticmethod
    def folder_counters(folder) -> dict:
        counters = {}
        for sub_folder in folder.Folders:
            counters[sub_folder.Name] = len(sub_folder.Items)
            if len(sub_folder.Folders) > 0:
                sub_counters = OutlookWrapper.folder_counters(folder)
                for key in sub_counters.keys():
                    counters[folder.Name + "/" + key] = sub_counters[key]
        return counters

    def sent_items_counter(self, date):
        count = 0
        sent_items = self.__sent_items.Items
        sent_email = sent_items.GetLast()
        while sent_email:
            if sent_email.SentOn.date() == date:
                count += 1
            elif sent_email.SentOn.date() < date:
                break
            sent_email = sent_items.GetPrevious()
        return count
