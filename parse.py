# -*- coding: utf-8 -*-
import os
import re
from enum import Enum
import sys
from os.path import isdir

Location = Enum('Location', 'Indoor Outdoor')
Gas = Enum('Gas', 'UNLEADED PLUS SUPREME')
CarWash = Enum('CarWash', 'Regular Deluxe Super')
Tender = Enum('Tender', 'Cash Credit Debit')
DIRECTORY_ANALYZED_FILENAME = 'ALREADY_ANALYZED'
FINALIZED_TRANSACTION_REGEX = re.compile(r'CUSTOMER\sTRANSACTION\s+(?P<txn_id>[0-9]+)\s+Finalized')
INDOOR_OUTDOOR_REGEX = re.compile(r'(?P<location>[Indoor|Outdoor]+)\s+tmnl(\s+)?:\s+(?P<terminal>[0-9]+)')
USER_SESSION_REGEX = re.compile(r'User\s+Session:\s+[0-9]+')
INITIAL_FUEL_PREPAY_REGEX = re.compile(r'Fuel\s+Prepay\s+Ref#(?P<ref_num>[0-9]+)\s+Pump\s+(?P<pump_num>[0-9]+)')
FINAL_FUEL_PREPAY_REGEX = re.compile(r'Original\s+Fuel\s+Prepay\s+Ref#(?P<ref_num>[0-9]+)')
FUEL_PREPAY_AMOUNT_REGEX = re.compile(r'FUEL\s+PREPAY\s+(?P<dollars>[0-9]+)\.(?P<cents>[0-9]+)')
TOTAL_DUE_REGEX = re.compile(r'TOTAL\s+DUE\s+[0-9]+\.[0-9]+')
BALANCE_DUE_REGEX = re.compile(r'BALANCE\s+DUE\s+[0-9]+\.[0-9]+')
INDOOR_TENDER_TYPE_REGEX = re.compile(r'(?P<tender>\w+)\s+.*')
DATE_TIME_REGEX = re.compile(r'(?P<month>[0-9]+)/(?P<day>[0-9]+)/(?P<yr>[0-9]+)\s+(?P<hr>[0-9]+):(?P<min>[0-9]+):(?P<sec>[0-9]+)')

class Txn(object):
    def __init__(self, id, date, time, amount, location, tender):
        """"""
        self.id = id
        self.date = date
        self.time = time
        self.amount = amount
        self.location = location
        self.tender = tender


class GasTxn(Txn):
    def __init__(self, id, date, time, amount, location, tender, volume, gas_type, pump_num, indoor_prepay=False, reference_num=None):
        """"""
        Txn.__init__(self, id, date, time, amount, location, tender)
        self.volume = volume
        self.gas_type = gas_type
        self.pump_num = pump_num
        self.indoor_prepay = indoor_prepay
        self.reference_num = reference_num



class CarWashTxn(Txn):
    """"""
    def __init__(self, id, date, time, amount, location, carwash_type, tender):
        Txn.__init__(self, id, date, time, amount, location, tender)
        self.carwash_type = carwash_type


def touch(fname, times=None):
    """"""
    with open(fname, 'a'):
        os.utime(fname, times)


def get_month_directories_to_analyze(dir_paths):
    """"""
    for p in dir_paths:
        if DIRECTORY_ANALYZED_FILENAME not in os.listdir(p):
            yield p


def mark_directory_as_analyzed(dir_path):
    """"""
    touch(fname=os.path.join(dir_path, DIRECTORY_ANALYZED_FILENAME))


def get_gas_transaction_from_line(txn_line, line_offsets, f):
    """
    txn_line + 1: Date and time
    txn_line + 2: Indoor or Outdoor
    if Outdoor:
        txn_line + 3: 'Outdoor tmnl: 1'
        txn_line + 6: '    UNLEADED PUR       52.50'
        txn_line + 8: 'Vol     12.592@     4.169'
        txn_line + 17: 'Card Type: Debit', 'Card Type: MASTERCARD'

    elif txn_line + 3 == 'User Session: [0-9]+':
        skip over this probably
        txn_line + 7: 'Fuel Prepay Ref#2740063 Pump 2'
        seek til 'TOTAL DUE             1.59'
        seek line + 2:
            if it == 'Cash                 18.41', ''
                it's cash


    elif indoor and txn_line + 3 == 'User Session: 6064' and txn_line + 6 == 'Original Fuel Prepay Ref#[0-9]+':
        it's a prepay
        create a mapping for the reference number with that transaction
        seek til TOTAL DUE:
            then + 2 == 'Cash                 18.41', 'Debit Card           33.69', '' <-- Probably cash
    elif indoor and txn_line + 6 == 'Original ..'
        pull it out of the map by reference number
        txn_line + 7: 'FUEL PREPAY\s+\-[0-9]+\.[0-9]+ (amount prepaid)
        txn_line + 8: 'UNLEADED PUR       17.00
        txn_line + 9: 'Ticket #923757      Pump 6'
        txn_line + 10: 'Vol      4.078@     4.169'
    """
    f.seek(line_offsets[txn_line])
    txn_line_offset = 0
    txn_id = re.match(FINALIZED_TRANSACTION_REGEX, f.readline()).group('txn_id')
    txn_line_offset += 1
    d_time_line = f.readline()
    dt_match = re.match(DATE_TIME_REGEX, d_time_line)
    date = '{0}/{1}/{2}'.format(dt_match.group('month'), dt_match.group('day'), dt_match.group('yr'))
    time = '{0}:{1}:{2}'.format(dt_match.group('hr'), dt_match.group('min'), dt_match.group('sec'))
    txn_line_offset += 1
    f.seek(line_offsets[txn_line + txn_line_offset])
    loc_tmnl_line = f.readline()
    loc_tmnl_match = re.match(INDOOR_OUTDOOR_REGEX, loc_tmnl_line)
    if loc_tmnl_match.group('location') == Location.Indoor.name:
        # Seek to the next line to test for user session
        txn_line_offset += 1
        f.seek(line_offsets[txn_line + txn_line_offset])
        u_ses_line = f.readline()
        user_session_match = re.match(USER_SESSION_REGEX, u_ses_line)
        if user_session_match:
            # Seek ahead to see if this is a fuel prepay. If TOTAL DUE is hit, it's not
            fuel_prepay_found, fuel_prepay_ref_match = False, None
            while not fuel_prepay_found:
                txn_line_offset += 1
                f.seek(line_offsets[txn_line + txn_line_offset])
                l = f.readline()
                fuel_prepay_ref_match = re.match(INITIAL_FUEL_PREPAY_REGEX, l)
                if fuel_prepay_ref_match:
                    fuel_prepay_found = True
                elif re.match(TOTAL_DUE_REGEX, l):
                    break
            if fuel_prepay_found:
                # We've got a prepay, create a new gas txn
                ref_num = fuel_prepay_ref_match.group('ref_num')
                pump_num = fuel_prepay_ref_match.group('pump_num')
                # Seek to the next line to get the amount
                txn_line_offset += 1
                f.seek(line_offsets[txn_line + txn_line_offset])
                f_pre_amt_line = f.readline()
                amt_match = re.match(FUEL_PREPAY_AMOUNT_REGEX, f_pre_amt_line)
                amount = float('{0}.{1}'.format(amt_match.group('dollars'), amt_match.group('cents')))
                # Seek until you reach the TOTAL DUE line
                total_due_line_found = False
                while not total_due_line_found:
                    txn_line_offset += 1
                    f.seek(line_offsets[txn_line + txn_line_offset])
                    l = f.readline()
                    if re.match(TOTAL_DUE_REGEX, l):
                        total_due_line_found = True
                # Seek ahead one line, if it's BALANCE DUE then it's cash and you're good
                txn_line_offset += 1
                f.seek(line_offsets[txn_line + txn_line_offset])
                next_line = f.readline()
                if re.match(BALANCE_DUE_REGEX, next_line):
                    tender = Tender.Cash.name
                else:
                    # Seek ahead one more line, and pull the word out
                    txn_line_offset += 1
                    f.seek(line_offsets[txn_line + txn_line_offset])
                    t_type_line = f.readline()

                    tender_type_match = re.match(INDOOR_TENDER_TYPE_REGEX, t_type_line)
                    if tender_type_match:
                        tender = tender_type_match.group('tender')
                    else:
                        tender = Tender.Cash.name
                prepay_txn = GasTxn(id=txn_id, date=date, time=time, amount=amount, location=Location.Indoor.name, tender=tender, volume=None, gas_type=None, pump_num=pump_num, indoor_prepay=True, reference_num=ref_num)
                return True, prepay_txn
            else:
                # It's not a prepay, return False, None
                return False, None
        else:
            # Seek to
            print 'Not a fuel prepay'
    else:
        # It's outdoor, this is a gas txn
        print 'OUT'
    return True, None


def get_gas_transactions_for_day(day_path):
    """"""
    gas_txns = []
    with open(day_path, 'r') as day:
        txn_lines, line_offsets = [], []
        offset = day.tell()
        for i, line in enumerate(day):
            line_offsets.append(offset)
            offset += len(line)
            m = re.match(FINALIZED_TRANSACTION_REGEX, line)
            if m:
                txn_lines.append(i)
        day.seek(0)
        for txn in txn_lines:
            is_gas_txn, gas_txn = get_gas_transaction_from_line(txn, line_offsets, day)




def main(args):
    """"""
    months_directory_path = './'
    if len(args) > 1:
        months_directory_path = args[1]
    month_directories = [os.path.join(months_directory_path, p) for p in os.listdir(months_directory_path) if isdir(os.path.join(months_directory_path, p))]
    dirs = get_month_directories_to_analyze(month_directories)
    for d in dirs:
        day_files = [os.path.join(d, p) for p in os.listdir(d) if os.path.splitext(p)[1] == '.txt']
        get_gas_transactions_for_day(day_files[0])


if __name__ == '__main__':
    main(sys.argv)