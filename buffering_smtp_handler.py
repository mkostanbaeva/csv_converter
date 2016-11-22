import logging, logging.handlers
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email import Encoders


class BufferingSMTPHandler(logging.handlers.BufferingHandler):
    def __init__(self, mailhost, mailport, timeout, fromaddr, toaddrs, subject, capacity,
                 logging_format, filename_sended):
        logging.handlers.BufferingHandler.__init__(self, capacity)
        self.mailhost = mailhost
        self.mailport = mailport
        self.timeout = timeout
        self.fromaddr = fromaddr
        self.toaddrs = toaddrs
        self.subject = subject
        self.formatter = logging_format
        self.setFormatter(logging.Formatter(logging_format))
        self.filename_sended = filename_sended

    def flush(self):
        if len(self.buffer) > 0:
            try:
                import smtplib
                port = self.mailport
                if not port:
                    port = smtplib.SMTP_PORT
                smtp = smtplib.SMTP(host=self.mailhost, port=port, timeout=self.timeout)
                if isinstance(self.toaddrs, list):  # If to addrs is a list, then join them as a string
                    toaddrs = ','.join(self.toaddrs)
                else:
                    toaddrs = self.toaddrs
                msg = "From: {}\r\nTo: {}\r\nSubject: {}\r\n\r\n".format(self.fromaddr, toaddrs,
                                                                         self.subject)
                for record in self.buffer:
                    s = self.format(record)
                    msg = msg + s + "\r\n"

                part = MIMEBase('application', "octet-stream")
                part.set_payload(open(self.filename_sended, "rb").read())
                Encoders.encode_base64(part)

                part.add_header('Content-Disposition', 'attachment; filename="self.filename_sended"')

                msg.attach(part)

                smtp.sendmail(self.fromaddr, self.toaddrs, msg)
                smtp.quit()
            except:
                self.handleError(None)  # no particular record
            self.buffer = []
