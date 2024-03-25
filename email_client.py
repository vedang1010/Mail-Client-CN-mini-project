from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QMessageBox, QDialog,QDialog, QTextBrowser
)
from PyQt5 import uic
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QTextCursor
import smtplib
import imaplib
import poplib
import email
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from time import sleep
from email.header import decode_header
import sys

from PyQt5.QtWidgets import QDialog, QTextBrowser
from PyQt5.QtGui import QTextCursor


class SentMailDialog(QDialog):
    
    def __init__(self, sent_emails, text_browser_sent):
        super(SentMailDialog, self).__init__()
        self.show()
        for email in sent_emails:
            text_browser_sent.append(f"From: {email.get('From')}")
            text_browser_sent.append(f"To: {email.get('To')}")
            text_browser_sent.append(f"Subject: {email.get('Subject')}")
            text_browser_sent.append(f"Date: {email.get('Date')}")
            text_browser_sent.append("\n")

        # Scroll to the bottom to show the latest emails
        text_browser_sent.moveCursor(QTextCursor.End)
    
class ReceiveMailDialog(QDialog):
    
    def __init__(self, emails, text_browser_received):
        super(ReceiveMailDialog, self).__init__()
        # uic.loadUi("receive_mail.ui", self)
        self.show()
        

        for email in emails:
            text_browser_received.append(f"From: {email.get('From')}")
            text_browser_received.append(f"To: {email.get('To')}")
            text_browser_received.append(f"Subject: {email.get('Subject')}")
            text_browser_received.append(f"Date: {email.get('Date')}")
            text_browser_received.append("\n")

        # Scroll to the bottom to show the latest emails
        text_browser_received.moveCursor(QTextCursor.End)
  

class MYGUI(QMainWindow):
    
    def __init__(self):
        super(MYGUI, self).__init__()
        uic.loadUi("mail.ui", self)
        self.show()
        self.sent_emails = [] 
        self.textBrowser_sent = QTextBrowser()
        self.text_browser_received = self.findChild(QTextBrowser, 'textBrowser_received')
        self.text_browser_sent = self.findChild(QTextBrowser, 'textBrowser_sent')
        self.pushButton.clicked.connect(self.login)
        self.pushButton_2.clicked.connect(self.attach_files)
        self.pushButton_3.clicked.connect(self.send_mail)
        self.pushButton_4.clicked.connect(self.receive_inbox)
        # self.pushButton_5 = QPushButton("Receive Sent Emails")
        self.pushButton_5.clicked.connect(self.receive_sent_emails)
        self.imap_server = "imap.gmail.com"
        self.imap = None   

        self.msg = MIMEMultipart()

    def login(self):
        try:
            selected_smtp_server = self.comboBox.currentText()
            self.server = smtplib.SMTP(selected_smtp_server, "587")
            self.server.ehlo()
            self.server.starttls()
            self.server.ehlo()
            self.server.login(self.lineEdit.text(), self.lineEdit_2.text())

            self.disable_login_controls()
            self.enable_send_controls()

        except Exception as e:
            self.show_error_message(f"Login failed: {str(e)}")

    def attach_files(self):
        options = QFileDialog.Options()
        filenames, _ = QFileDialog.getOpenFileNames(
            self, "Open File", "", "All Files (*.*)", options=options)
        if filenames:
            for filename in filenames:
                try:
                    with open(filename, 'rb') as attachment:
                        file_content = attachment.read()

                    base_part = MIMEBase('application', 'octet-stream')
                    base_part.set_payload(file_content)
                    encoders.encode_base64(base_part)
                    base_part.add_header(
                        "Content-Disposition", f"attachment; filename={filename}")

                    self.msg.attach(base_part)

                    self.update_attachments_label(filename)

                except Exception as e:
                    self.show_error_message(f"Error attaching file {filename}: {str(e)}")

    def send_mail(self):
        try:
            self.compose_mail()
            self.send_mail_message()

            self.show_info_message("Mail sent!")

        except Exception as e:
            self.show_error_message(f"Sending Mail Failed: {str(e)}")

    def decode_subject(subject):
        decoded_parts = decode_header(subject)
        decoded_subject = ''

        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                try:
                    decoded_subject += part.decode(encoding or 'utf-8', errors='strict')
                except UnicodeDecodeError:
                    decoded_subject += part.decode('latin-1', errors='replace')
            else:
                decoded_subject += str(part)

        return decoded_subject

    def fetch_and_print_messages(self, imap, mailbox):
        try:
            _, msgnums = imap.search(None, 'ALL')
            latest_msgnums = msgnums[0].split()[-10:]

            for msgnum in latest_msgnums:
                _, data = imap.fetch(msgnum, "(RFC822)")
                message = email.message_from_bytes(data[0][1])
                print(f"Message Number: {msgnum}")

                self.print_email_details(message)

        except Exception as e:
            self.show_error_message(f"Error fetching messages: {str(e)}")

    def receive_inbox(self):
        try:
            # Get the selected protocol from the combo box
            selected_protocol = self.comboBox_2.currentText()

            if selected_protocol == "imap.gmail.com":
                self.receive_imap()
            elif selected_protocol == "outlook.office365.com":
                self.receive_pop3()
            else:
                self.show_error_message("Invalid protocol selected")

        except Exception as e:
            self.show_error_message(f"Error receiving inbox: {str(e)}")

    def receive_imap(self):
        try:
            imap_server = self.comboBox_2.currentText()
            self.imap = imaplib.IMAP4_SSL(imap_server)

            email_address = self.lineEdit.text()
            password =self.lineEdit_2.text()
            self.imap.login(email_address, password)

            self.imap.select("inbox")

            sleep(2)

            print("Messages from INBOX:")
            emails = self.fetch_and_get_messages("INBOX")

            self.imap.close()
            self.imap.logout()

            receive_dialog = ReceiveMailDialog(emails, self.text_browser_received)

        except Exception as e:
            self.show_error_message(f"Error receiving inbox: {str(e)}")

    def receive_pop3(self):
        try:
            # Assuming POP3 server uses SSL, if not, use 'poplib.POP3' instead.
            pop_server = self.comboBox_2.currentText()
            self.pop3 = poplib.POP3_SSL(pop_server)

            # Login to your account
            email_address = "khedekarvp21.comp@coeptech.ac.in"
            password = "Vedu@123"
            self.pop3.user(email_address)
            self.pop3.pass_(password)

            # Get the number of messages in the mailbox
            num_messages = len(self.pop3.list()[1])

            # Fetch and print messages from the inbox
            print("Messages from INBOX:")
            emails = self.fetch_and_get_messages(num_messages)

            # Close the mailbox
            self.pop3.quit()

            # Open the dialog to display received emails
            receive_dialog = ReceiveMailDialog(emails, self.text_browser_received)

        except Exception as e:
            self.show_error_message(f"Error receiving inbox: {str(e)}")
       
    def fetch_and_get_messages(self, mailbox):
        emails = []
        _, msgnums = self.imap.search(None, 'ALL')
        latest_msgnums = msgnums[0].split()[-10:]

        for msgnum in latest_msgnums:
            _, data = self.imap.fetch(msgnum, "(RFC822)")
            message = email.message_from_bytes(data[0][1])
            emails.append(message)
            self.print_email_details(message)

        return emails
    
    def disable_login_controls(self):
        self.lineEdit.setEnabled(False)
        self.lineEdit_2.setEnabled(False)
        self.comboBox.setEnabled(False)
        self.lineEdit_4.setEnabled(False)
        self.comboBox_2.setEnabled(False)
        self.pushButton.setEnabled(False)

    def enable_send_controls(self):
        self.lineEdit_5.setEnabled(True)
        self.lineEdit_6.setEnabled(True)
        self.textEdit.setEnabled(True)
        self.pushButton_2.setEnabled(True)
        self.pushButton_3.setEnabled(True)
        self.pushButton_4.setEnabled(True)
        self.pushButton_5.setEnabled(True)

    def update_attachments_label(self, filename):
        if not self.label_8.text().endswith(":"):
            self.label_8.setText(self.label_8.text() + ", ")
        self.label_8.setText(self.label_8.text() + filename)

    def compose_mail(self):
        self.msg['From'] = "Vedang Khedekar"
        self.msg['To'] = self.lineEdit_5.text()
        self.msg['Subject'] = self.lineEdit_6.text()
        self.msg.attach(MIMEText(self.textEdit.toPlainText(), 'plain'))

    def send_mail_message(self):
        text = self.msg.as_string()
        self.server.sendmail(self.lineEdit.text(), self.lineEdit_5.text(), text)

    def print_email_details(self, message):
        try:
            subject = message["Subject"]
            body = ""

            if message.is_multipart():
                for part in message.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True)
            else:
                body = message.get_payload(decode=True)

            print(f"From: {decode_header(message.get('From'))[0][0]}")
            print(f"To: {decode_header(message.get('To'))[0][0]}")
            print(f"BCC: {decode_header(message.get('BCC'))[0][0]}")
            print(f"Date: {decode_header(message.get('Date'))[0][0]}")
            print(f"Subject: {decode_header(subject)[0][0]}")
            print(f"Body: {body.decode_header('utf-8', errors='replace')}")

        except Exception as e:
            print(f"Error printing email details: {str(e)}")
            
    def show_info_message(self, message):
        message_box = QMessageBox()
        message_box.setText(message)
        message_box.exec_()

    def show_error_message(self, message):
        QMessageBox.critical(self, "Error", message)

    def receive_sent_emails(self):
        try:
            imap_server = self.comboBox_2.currentText()
            self.imap = imaplib.IMAP4_SSL(imap_server)

            # Login to your account
            # email_address = "vedangkhedekar07@gmail.com"
            # password = "moml vnrc nere uqri"
            email_address = self.lineEdit.text()
            password =self.lineEdit_2.text()
            self.imap.login(email_address, password)
            
            self.imap.select('"[Gmail]/Sent Mail"')

            print("Last 10 Sent Messages:")
            sent_emails = self.fetch_and_get_messages("Sent Mail")
            sent_dialog = SentMailDialog(sent_emails, self.text_browser_sent)
            # sent_dialog.exec_()
            self.imap.close()
            self.imap.logout()

            # Open the dialog to display sent emails
          

        except Exception as e:
            self.show_error_message(f"Error fetching sent emails: {str(e)}")

    # def fetch_and_get_messages(self, mailbox):
    #     emails = []
    #     _, msgnums = self.imap.search(None, 'ALL')
    #     latest_msgnums = msgnums[0].split()[-10:]

    #     for msgnum in latest_msgnums:
    #         _, data = self.imap.fetch(msgnum, "(RFC822)")
    #         message = email.message_from_bytes(data[0][1])
    #         emails.append(message)

    #     return emails

    
app = QApplication([])
window = MYGUI()
sys.exit(app.exec_())
