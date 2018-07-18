#!/usr/bin/env python
from __future__ import print_function
import vobject
import os
import sys
import csv


def write_record_to_card(record, vcard_path):
    """ Writes the CSV record as a vcard to the given dir. """
    first = record[1]
    middle = record[2]
    last = record[3]
    suffix = record[4]
    company = record[5]
    dept = record[6]
    title = record[7]
    bus_add_1 = record[8]
    bus_add_2 = record[9]
    # bus_add_3 = record[10]
    bus_city = record[11]
    bus_state = record[12]
    bus_postal = record[13]
    bus_region = record[14]
    home_add_1 = record[15]
    home_add_2 = record[16]
    # home_add_3 = record[17]
    home_city = record[18]
    home_state = record[19]
    home_postal = record[20]
    home_region = record[21]
    other_add_1 = record[22]
    other_add_2 = record[23]
    # other_add_3 = record[24]
    other_city = record[25]
    other_state = record[26]
    other_postal = record[27]
    other_region = record[28]
    fax = record[29]
    bus_phone = record[30]
    bus_phone_2 = record[31]
    home_phone = record[36]
    mobile = record[39]
    email = record[46]

    v = vobject.vCard()

    # name
    v.add('n').value = vobject.vcard.Name(
        family=last, given=first,
        additional=middle, suffix=suffix,
    )
    v.add('fn').value = "{} {}".format(first, last)

    # title
    if title is not '':
        v.add('title').value = title

    # company
    if company is not '':
        v.add('org')
        v.org.value = [company, dept]

    # business address
    if bus_add_1 is not '':
        vadd = v.add('adr')
        vadd.value = vobject.vcard.Address(
            street=bus_add_1 + ' ' + bus_add_2, city=bus_city,
            region=bus_state, country=bus_region, code=bus_postal,
        )
        vadd.type_param = 'WORK'

    # home address
    if home_add_1 is not '':
        vadd = v.add('adr')
        vadd.value = vobject.vcard.Address(
            street=home_add_1 + ' ' + home_add_2, city=home_city,
            region=home_state, country=home_region, code=home_postal,
        )
        vadd.type_param = 'HOME'

    # other address
    if other_add_1 is not '':
        vadd = v.add('adr')
        vadd.value = vobject.vcard.Address(
            street=other_add_1 + ' ' + other_add_2, city=other_city,
            region=other_state, country=other_region, code=other_postal,
        )
        vadd.type_param = 'OTHER'

    # phone numbers
    if bus_phone is not '':
        vphone = v.add('tel')
        vphone.value = bus_phone
        vphone.type_param = 'WORK'
    if bus_phone_2 is not '':
        vphone = v.add('tel')
        vphone.value = bus_phone_2
        vphone.type_param = 'WORK2'
    if home_phone is not '':
        vphone = v.add('tel')
        vphone.value = home_phone
        vphone.type_param = 'HOME'
    if mobile is not '':
        vphone = v.add('tel')
        vphone.value = mobile
        vphone.type_param = 'CELL'
    if fax is not '':
        vphone = v.add('tel')
        vphone.value = fax
        vphone.type_param = 'FAX'

    # email
    if email is not '':
        v.add('email').value = email

    vfilename = "{}_{}.vcard".format(last, first)
    file_path = os.path.join(vcard_path, vfilename)
    with open(file_path, 'w') as vcard_file:
        v.serialize(vcard_file, lineLength=180)


def get_records_from_csv(csv_filepath):
    """ Gets a generator of records from a CSV file. """
    with open(csv_filepath, 'rb') as csvfile:
        reader = csv.reader(csvfile)
        for record in reader:
            yield record


def run(csv_filepath, vcard_output_dir):
    for record in get_records_from_csv(csv_filepath):
        write_record_to_card(record, vcard_output_dir)


def help():
    """ Prints the help. """
    print("Usage: csvtovcard.py INPUT_FILE OUTPUT_DIR")


if __name__ == '__main__':
    args = sys.argv
    if len(args) != 3:
        help()
        exit(1)

    csv_file = args[1]
    if not os.path.exists(csv_file):
        print(
            csv_file,
            "does not exist!",
            file=sys.stderr
        )
        exit(1)

    output_dir = args[2]
    if not os.path.isdir(output_dir):
        os.mkdir(output_dir)

    run(csv_file, output_dir)

    exit(0)
