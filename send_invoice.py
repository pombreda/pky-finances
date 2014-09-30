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
from datetime import datetime

DETAILS_TEMPLATE = u"""
Saaja: Polyteknikkojen Kuoron kannatusyhdistys ry
Pankkiyhteys: Nordea
Tilinumero: FI14 1112 3000 3084 34
Viitenumero: %(viitenro)s
Summa: %(summa)s
Eräpäivä: %(eräpäivä)s
"""

FOOTER = u"""
Terveisin,
  Markus Lehtonen
  PKY
"""

def compose_email(message, row_data):
    """Compose email text"""
    return message + '\n' + DETAILS_TEMPLATE % row_data + FOOTER

def utf8_reader(input_stream):
    """Wrapper for reading UTF8-encoded csv files"""
    reader = csv.reader(input_stream)
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

def parse_args(argv):
    """Parse command line arguments"""
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--dry-run', action='store_true',
                        help='Do everything but send email')
    parser.add_argument('--cc',
                        help='Carbon copy to this email')
    parser.add_argument('-m', '--message',
                        help='Greeting message')
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

    message = args.message or "Hei,\n\nOhessa lasku.\n"
    message = message.decode('string_escape').decode('utf-8')

    with open(args.csv, 'r') as fobj:
        reader = utf8_reader(fobj)
        print "CSV file time stamp:", reader.next()[0]
        header_row = reader.next()

        all_data = []
        for row in reader:
            all_data.append(dict(zip([val.lower() for val in header_row], row)))

    if args.range:
        send_data = [row for row in all_data if
                        in_range(row[u'nro'], args.range)]
    else:
        date = args.date or datetime.now().date()
        send_data = [row for row in all_data if
                        row[u'nro'] and std_date(row[u'pvm']) == date]

    for ind, row in enumerate(send_data, 1):
        info = "INVOICE #%d to <%s>" % (ind, row['email'])
        print "\n==== " + info + " " + "="*(76-5-len(info))
        print compose_email(message, row)

if __name__ == '__main__':
    main()
