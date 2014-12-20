# -*- coding: utf-8 -*-
import os
import re
from enum import Enum
import sys
from os.path import isdir

Location = Enum('Location', 'Indoor Outdoor')
Gas = Enum('Gas', 'Regular Premium Supreme')
CarWash = Enum('CarWash', 'Regular Deluxe Super')
DIRECTORY_ANALYZED_FILENAME = 'ALREADY_ANALYZED'
FINALIZED_TRANSACTION_REGEX = re.compile(r'CUSTOMER\sTRANSACTION\s+(?P<txn_id>[0-9]+)\s+Finalized')


class Txn(object):
    def __init__(self, id, datetime, amount, location):
        """"""
        self.id = id
        self.datetime = datetime
        self.amount = amount
        self.location = location


class GasTxn(Txn):
    def __init__(self, id, datetime, amount, location, volume, type):
        """"""
        Txn.__init__(self, id, datetime, amount, location)
        self.volume = volume
        self.type = type


class CarWashTxn(Txn):
    """"""
    def __init__(self, id, datetime, amount, location, type):
        Txn.__init__(self, id, datetime, amount, location)
        self.type = type


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


def get_gas_transactions_for_day(day_path):
    """"""
    print day_path
    with open(day_path, 'r') as day:
        for line in day:
            l = line.strip()
            m = re.match(FINALIZED_TRANSACTION_REGEX, l)
            if m:
                print m.group('txn_id')


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