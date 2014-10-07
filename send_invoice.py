#!/usr/bin/python
# vim:fileencoding=utf-8:et:ts=4:sw=4:sts=4
#
# Copyright (C) 2014 Markus Lehtonen <knaeaepae@gmail.com>
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
# MA 02110-1301, USA.
"""Helper for sending PKY invoices"""

import argparse
import email.charset
import csv
import os
import re
import smtplib
import string
import sys
from ConfigParser import ConfigParser
from datetime import datetime
from email.header import Header, decode_header
from email.mime.text import MIMEText


DETAILS_TEMPLATE = u"""
Selite: %(selite)s
Saaja: Polyteknikkojen Kuoron kannatusyhdistys ry
Pankkiyhteys: Nordea
Tilinumero: FI14 1112 3000 3084 34
Viitenumero: %(viitenro)s
Summa: %(summa)s
Eräpäivä: %(eräpäivä)s
"""

FOOTER = u"""

Parhain terveisin,
  Markus Lehtonen
  PKY
"""

def compose_email(headers, greeting_msg, row_data):
    """Compose email text"""
    msg = MIMEText(greeting_msg + '\n' + DETAILS_TEMPLATE % row_data + FOOTER,
                   _charset='utf-8')
    for key, val in headers.iteritems():
        msg[key.capitalize()] = val
    return msg

def pprint_email(msg):
    """Pretty print email"""
    for key, val in msg.items():
        # Filter out uninteresting parts
        if key.lower() not in ['content-type', 'mime-version',
                               'content-transfer-encoding']:
            print '%s: %s' % (key, unicode(val))
    print ""
    print msg.get_payload(decode=True)


def utf8_reader(input_stream, dialect=None):
    """Wrapper for reading UTF8-encoded csv files"""
    reader = csv.reader(input_stream, dialect=dialect)
    for row in reader:
        yield [unicode(cell, 'utf-8') for cell in row]

def in_range(val, range_str):
    """Is value in range

    >>> in_range('1', '1')
    True
    >>> in_range('1', '2')
    False
    >>> in_range('6', '1,3-8,10')
    True
    >>> in_range('9', '1,3-8,10')
    False
    """
    val = int(val)
    ranges = range_str.split(',')
    for ran in ranges:
        split = ran.split('-', 1)
        if len(split) == 1:
            if val == int(split[0]):
                return True
        else:
            if val >= int(split[0]) and val <= int(split[1]):
                return True
    return False

def to_u(text):
    """Convert text to unicode, assumes UTF-8 for str input"""
    if isinstance(text, str):
        return unicode(text, 'utf-8')
    else:
        return unicode(text)

def utf8_header(text):
    """Email header wih UTF-8 encoding"""
    # Convert text to unicode (assume we're using UTF-8)
    return Header(to_u(text), 'utf-8')

def utf8_address_header(addr):
    """Create an internationalized header from name, email tuple"""
    if isinstance(addr, tuple) or isinstance(addr, basestring):
        addr = [addr]

    header = utf8_header('')
    for address in addr:
        if isinstance(address, basestring):
            name, email = split_email_address(address)
        else:
            name, email = address
        if str(header):
            header.append(u', ')
        if name:
            header.append(to_u(name))
            header.append(to_u(' <%s>' % email))
        else:
            header.append(to_u(email))
    return header

def split_email_address(text):
    """Split name and address out of an email address

    >>> split_email_address('foo@bar.com')
    ('', 'foo@bar.com')
    >>> split_email_address('Foo Bar foo@bar.com')
    ('Foo Bar', 'foo@bar.com')
    >>> split_email_address('  "Foo Bar" <foo@bar.com>, ')
    ('Foo Bar', 'foo@bar.com')
    """
    split = text.strip().rsplit(None, 1)
    email_re = r'.*?([^<%s]\S*@\S+[a-zA-Z])' % string.whitespace
    match = re.match(email_re, split[-1])
    if not match:
        raise Exception("Invalid email address: '%s'" % text)
    email = match.group(1)

    name = ''
    if len(split) > 1:
        non_letter = string.punctuation + string.whitespace
        name_re = r'.*?([^%s].*[^%s])' % (non_letter, non_letter)
        match = re.match(name_re, split[0])
        if match:
            name = match.group(1)
    return (name, email)

def std_date(string):
    """Convert string to date"""
    return datetime.strptime(string, '%d.%m.%Y').date()

def ask_value(question, default=None, choices=None):
    """Ask user input"""
    choice_str = ' (%s)' % '/'.join(choices) if choices else ''
    default_str = ' [%s]' % default if default is not None else ''
    while True:

        val = raw_input(question + '%s:%s ' % (choice_str, default_str))
        if val:
            if not choices or val in choices:
                return val
        elif default is not None:
            return default

def parse_config(path):
    """Read config file"""
    defaults = {'smtp-server': '',
                'from': '',
                'subject-prefix': ''}
    parser = ConfigParser(defaults)
    parser.add_section('general')

    filepath = os.path.join(path, 'send_invoice.conf')
    confs = parser.read(filepath)
    if confs:
        print "Read config file %s" % confs
    else:
        print "Did not find config file %s" % filepath

    # Only use one section, i.e. 'general'
    return dict(parser.items('general'))


def parse_args(argv):
    """Parse command line arguments"""
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--dry-run', action='store_true',
                        help='Do everything but send email')
    parser.add_argument('--from', dest='sender', type=split_email_address,
                        help="Sender's email")
    parser.add_argument('--cc', action='append', default=[],
                        type=split_email_address,
                        help='Carbon copy to this email')
    parser.add_argument('--bcc', action='append', default=[],
                        type=split_email_address,
                        help='Blind (hidden) carbon copy to this email')
    parser.add_argument('--smtp-server', help="Address of the SMTP server")
    parser.add_argument('-m', '--message',
                        help='Greeting message, used for all invoices')
    parser.add_argument('--subject',
                        help="Messgae subject, used for all invoices")
    parser.add_argument('--subject-prefix', metavar='PREFIX', default='',
                        help='Prefix all email subjects with %(metavar)s')
    parser.add_argument('-G', '--group-by', metavar='COLUMN', default='viite',
                        help='Mass-send invoices with the same value of COLUMN')
    group = parser.add_mutually_exclusive_group()
    group.add_argument('-D', '--date', type=std_date,
                       help='Send invoices having this date')
    group.add_argument('-I', '--index',
                       help='Send invoices having these index numbers')

    parser.add_argument('csv',
                        help='CSV file containing invoice entries')
    return parser.parse_args(argv[1:])

def main(argv=None):
    """Script entry point"""

    print "Welcome to PKY invoice sender!"

    args = parse_args(argv)
    config = parse_config(os.path.dirname(argv[0]))

    # Change email header encoding to QP for easier readability of raw data
    email.charset.add_charset('utf-8', email.charset.QP, email.charset.QP)

    with open(args.csv, 'r') as fobj:
        dialect = csv.Sniffer().sniff(fobj.read(512))
        fobj.seek(0)
        reader = utf8_reader(fobj, dialect)
        print "CSV file time stamp:", reader.next()[0]
        header_row = reader.next()

        all_data = []
        for row in reader:
            all_data.append(dict(zip([val.lower() for val in header_row], row)))

    if args.index:
        send_data = [row for row in all_data if
                        row[u'nro'] and in_range(row[u'nro'], args.index)]
    else:
        date = args.date or datetime.now().date()
        send_data = [row for row in all_data if
                        row[u'nro'] and std_date(row[u'pvm']) == date]

    if not send_data:
        print "No invoices to send, exiting"
        return 0

    # Get SMTP server
    if args.smtp_server:
        smtp_server = args.smtp_server
    else:
        smtp_server = config['smtp-server'] or ask_value('SMTP server')

    # Get sender email
    if args.sender:
        sender = args.sender
    else:
        if 'EMAIL' in os.environ:
            sender = os.environ['EMAIL']
        sender = split_email_address(config['from'] or
                                     ask_value('From', default=sender))

    server = smtplib.SMTP(smtp_server)

    # Get subject prefix
    subject_prefix = args.subject_prefix or config['subject-prefix']
    try:
        # Group data
        if args.group_by:
            groups = set([row[args.group_by] for row in send_data])
            grouped_data = []
            for group in groups:
                grouped_data.append([row for row in send_data if
                            row[args.group_by] == group])
            assert sum([len(grp) for grp in grouped_data]) == len(send_data)
        else:
            grouped_data = [[row] for row in send_data]

        # Send grouped emails
        for g_ind, group in enumerate(grouped_data, 1):
            if len(group) > 1:
                info_header = "#%d: INVOICE GROUP " % g_ind
                info_msg = "%s: %s\n" % (args.group_by.upper(),
                                         group[0][args.group_by])
            else:
                info_header = "#%d: SIGNLE INVOICE " % g_ind
                info_msg = 'EMAIL: %(email)s\nVIITE: %(viite)s\n' \
                           'SUMMA:%(summa)s\n' % group[0]
            print "\n==== " + info_header + " " + "="*(76-5-len(info_header))
            print info_msg

            # Get greeting message
            if args.message:
                message = args.message
            else:
                message = "Hei,\n\nOhessa lasku."
                message = ask_value('Greeting message', default=message)
            message = message.decode('string_escape').decode('utf-8')

            # Email headers
            headers = {'from': utf8_address_header(sender)}

            subject_prefix = subject_prefix + ' ' if subject_prefix else ''
            if args.subject:
                headers['subject'] = utf8_header(subject_prefix + args.subject)
            else:
                headers['subject'] = utf8_header(subject_prefix +
                                                 ask_value('Subject'))
            if args.cc:
                headers['cc'] = utf8_address_header(args.cc)
            if args.bcc:
                headers['bcc'] = utf8_address_header(args.bcc)

            # Ask for confirmation
            headers['to'] = utf8_address_header(group[0]['email'])
            example = compose_email(headers, message, group[0])
            print '\n' + '-' * 79
            pprint_email(example)
            print '-' * 79 + '\n'
            recipients = ['<%s>' % split_email_address(row['email'])[1] for
                            row in group]
            proceed = ask_value("Send an email like above to %s" %
                                ', '.join(recipients), choices=['n', 'y'])
            if proceed == 'y':
                for row in group:
                    to_name, to_email = split_email_address(row['email'])
                    recipients = [to_email] + \
                                 [cc[1] for cc in args.cc] + \
                                 [bcc[1] for bcc in args.bcc]
                    recipients = [row['email']] + args.cc + args.bcc
                    headers['to'] = utf8_address_header((to_name, to_email))
                    msg = compose_email(headers, message, row)
                    print "Sending email to <%s>..." % recipients[0]
                    if not args.dry_run:
                        rsp = server.sendmail(sender,
                                        recipients, msg.as_string(),
                                        rcpt_options=['NOTIFY=FAILURE,DELAY'])
                        if rsp:
                            print "Mail delivery failed: %s" % rsp
            else:
                print "Did not send!"

    finally:
        server.quit()

    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
