#!/usr/bin/python
# -*- coding: UTF-8 -*-

# Based on the work of
# https://github.com/evanj/hsbc2ofx/blob/master/hsbc2ofx.py

# Written by Robert Myhren (myhren@gmail.com)
# Tool to convert csvfiles from Arendal & Omegn Sparekasse to OFX files

import csv
import math
import sys
import time

# Substitute your bank ID and account ID here:
BANKNAME = "Arendal & Omegn Sparekasse"
BANKID = "SPAREKASSEN"
ACCOUNTID = "ROBERT"

# These strings are removed from the "name" field to make it more legible
REMOVE_NAME_PREFIXES = (
    'Avtalegiro til ',
    'Overført til ',
    'Varekjøp ',
    'Giro - Melding - ',
)

# According to the Quickbooks document, <NAME> field is this many chars
NAME_LENGTH = 32
# Maximum characters in the <MEMO> field
MEMO_LENGTH = 255

# Change amount to corrct format
def amount_to_float(amount_string):
  if amount_string == '':
    return None

  amount_string = amount_string.replace('.', '')
  amount_string = amount_string.replace(',', '.')
  return float(amount_string)

def main():
    input_path = sys.argv[1]
    dialect = csv.Sniffer().sniff(open(input_path).read())
    f = open(input_path)
    lines = f.read().splitlines()
    f.close()
    reader = csv.reader(lines[1:], csv.excel, delimiter=';', quotechar='"')

    end_date_now = time.strftime('%Y%m%d', time.gmtime())

    print '''\nOFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE
<OFX>
    <SIGNONMSGSRSV1>
        <SONRS>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <DTSERVER>%s</DTSERVER>
            <LANGUAGE>NO</LANGUAGE>
            <FI>
                <ORG>%s</ORG>
            </FI>
        </SONRS>
    </SIGNONMSGSRSV1>
    <BANKMSGSRSV1>
        <STMTTRNRS>
            <TRNUID>1</TRNUID>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <STMTRS>
            <CURDEF>NOK</CURDEF>
            <BANKACCTFROM>
                <BANKID>%s</BANKID>
                <ACCTID>%s</ACCTID>
                <ACCTTYPE>CHECKING</ACCTTYPE>
            </BANKACCTFROM>
            <BANKTRANLIST>
                <DTSTART>19700101</DTSTART>
                <DTEND>%s</DTEND>\n''' % (end_date_now, BANKNAME ,BANKID, ACCOUNTID, end_date_now)

    for row in reader:
        date    = row[1]
        desc    = row[2]
        amount  = row[3]
        balance = row[4]

        amount = amount_to_float(amount)

        day, month, year = date.split('.')
        day = int(day)
        month = int(month)
        year = int(year)

            # Fake an ID with date + amount
        fitid = '%02d%02d%02d%.2f' % (year, month, day, math.fabs(amount))

            # remove useless prefixes then truncate
        short_desc = desc
        for prefix in REMOVE_NAME_PREFIXES:
            if short_desc.startswith(prefix):
                short_desc = short_desc[len(prefix):]
        short_desc = short_desc[:NAME_LENGTH]

        indent = '                '
        print '%s<STMTTRN>' % (indent)
        print '%s    <TRNTYPE>%s</TRNTYPE>' % (indent, "CHECK")
        print '%s    <DTPOSTED>%02d%02d%02d</DTPOSTED>' % (indent, year, month, day)
        print '%s    <TRNAMT>%.2f</TRNAMT>' % (indent, amount)
        # if ofx_type == 'CHECK':
        #   print '%s    <CHECKNUM>%s</CHECKNUM>' % (indent, customer_reference)
        print '%s    <FITID>%s</FITID>' % (indent, fitid)
        print '%s    <NAME>%s</NAME>' % (indent, short_desc)
        print '%s    <MEMO>%s</MEMO>' % (indent, desc[:MEMO_LENGTH])
        print '%s</STMTTRN>' % (indent)

    print '''            </BANKTRANLIST>
                <LEDGERBAL>
                    <BALAMT>17752.42</BALAMT>
                    <DTASOF>20130930</DTASOF>
                </LEDGERBAL>
            </STMTRS>
        </STMTTRNRS>
    </BANKMSGSRSV1>
</OFX>
'''


if __name__ == '__main__':
  if len(sys.argv) != 2:
    sys.stderr.write('sparekassen2ofx.py (input CSV)\n')
    sys.exit(1)
  sys.exit(main())
