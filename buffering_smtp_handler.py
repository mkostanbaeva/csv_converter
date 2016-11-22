import logging, logging.handlers
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email import Encoders
from email.mime.text import MIMEText


class BufferingSMTPHandler(logging.handlers.BufferingHandler):
    def __init__(self, mailhost, mailport, timeout, fromaddr, toaddrs, subject, capacity,
                 logging_format):
        logging.handlers.BufferingHandler.__init__(self, capacity)
        self.mailhost = mailhost
        self.mailport = mailport
        self.timeout = timeout
        self.fromaddr = fromaddr
        self.toaddrs = toaddrs
        self.subject = subject
        self.formatter = logging_format
        self.setFormatter(logging.Formatter(logging_format))

    def flush(self, filename=None):
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

                msg = MIMEMultipart()
                msg['Subject'] = self.subject
                msg['From'] = self.fromaddr
                msg['To'] = toaddrs

                text = ''
                for record in self.buffer:
                    s = self.format(record)
                    text = text + s + "\r\n"
                body = MIMEText(text, 'plain')
                msg.attach(body)

                if filename is not None:
                    part = MIMEBase('application', "octet-stream")
                    part.set_payload(open(filename, "rb").read())
                    Encoders.encode_base64(part)

                    part.add_header('Content-Disposition', 'attachment; filename="filename"')

                    msg.attach(part)

                smtp.sendmail(self.fromaddr, self.toaddrs, msg.as_string())
                smtp.quit()
            except:
                self.handleError(None)  # no particular record
            self.buffer = []
