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
import csv
import os
import smtplib
import sys
from datetime import datetime
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

def compose_email(sender, recipient, subject, message, row_data):
    """Compose email text"""
    msg = MIMEText(message + '\n' + DETAILS_TEMPLATE % row_data + FOOTER,
                   _charset='utf-8')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = recipient
    return msg

def pprint_email(msg):
    """Pretty print email"""
    print "From:", msg['From']
    print "To:", msg['To']
    print "Subject:", msg['Subject']
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

def parse_args(argv):
    """Parse command line arguments"""
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--dry-run', action='store_true',
                        help='Do everything but send email')
    parser.add_argument('--from', dest='sender', help="Sender's email")
    parser.add_argument('--cc',
                        help='Carbon copy to this email')
    parser.add_argument('--smtp-server', help="Address of the SMTP server")
    parser.add_argument('-m', '--message',
                        help='Greeting message, used for all invoices')
    parser.add_argument('--subject',
                        help="Messgae subject, used for all invoices")
    parser.add_argument('-G', '--group-by', metavar='COLUMN', default='selite',
                        help='Mass-send invoices with the same value of COLUMN')
    group = parser.add_mutually_exclusive_group()
    group.add_argument('-D', '--date', type=std_date,
                       help='Send invoices having this date')
    group.add_argument('-R', '--range',
                       help='Send invoices having these index numbers')

    parser.add_argument('csv',
                        help='CSV file containing invoice entries')
    return parser.parse_args(argv)

def main(argv=None):
    """Script entry point"""

    print "Welcome to PKY invoice sender!"

    args = parse_args(argv)

    with open(args.csv, 'r') as fobj:
        dialect = csv.Sniffer().sniff(fobj.read(512))
        fobj.seek(0)
        reader = utf8_reader(fobj, dialect)
        print "CSV file time stamp:", reader.next()[0]
        header_row = reader.next()

        all_data = []
        for row in reader:
            all_data.append(dict(zip([val.lower() for val in header_row], row)))

    if args.range:
        send_data = [row for row in all_data if
                        row[u'nro'] and in_range(row[u'nro'], args.range)]
    else:
        date = args.date or datetime.now().date()
        send_data = [row for row in all_data if
                        row[u'nro'] and std_date(row[u'pvm']) == date]

    if not send_data:
        print "No invoices to send, exiting"
        return 0

    # Get SMTP server
    smtp_server = args.smtp_server or ask_value('SMTP server')

    # Get sender email
    if args.sender:
        sender = args.sender
    else:
        if 'EMAIL' in os.environ:
            sender = os.environ['EMAIL']
        sender = ask_value('From', default=sender)

    server = smtplib.SMTP(smtp_server)

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
                info_msg = 'EMAIL: %(email)s\nSELITE: %(selite)s\n' \
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

            # Get subject
            if args.subject:
                subject = args.subject
            else:
                subject = ask_value('Subject')

            # Ask for confirmation
            example = compose_email(sender, group[0]['email'], subject, message,
                                    group[0])
            print '\n' + '-' * 79
            pprint_email(example)
            print '-' * 79 + '\n'
            recipients = ['<%s>' % row['email'] for row in group]
            proceed = ask_value("Send an email like above to %s" %
                                ', '.join(recipients), choices=['n', 'y'])
            if proceed == 'y':
                for row in group:
                    recipient = row['email']
                    msg = compose_email(sender, recipient, subject, message,
                                        row)
                    print "Sending email to <%s>..." % recipient
                    server.sendmail(sender, [recipient], msg.as_string())
            else:
                print "Did not send!"

    finally:
        server.quit()

    return 0


if __name__ == '__main__':
    sys.exit(main())
