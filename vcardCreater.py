# EStroev
import os
import csv
import openpyxl
import argparse


def xlsx_opener(inFile, title):
    inWB = openpyxl.load_workbook(inFile, read_only=True)
    print(f'[+] Open {inFile}:{title}')
    ws = inWB[title]

    return ws


def load_data(ws):
    contacts = list()
    for cell in ws.iter_rows(row_offset=1):
        contacts.append(
            {
                'name': cell[0].value,
                'surname': cell[1].value,
                'phone1': cell[2].value,
                'phone2': cell[3].value,
                'workEmail': cell[4].value,
                'personalEmail': cell[5].value,
                'company': cell[6].value,
                'title': cell[7].value
            }
        )

    return contacts


def xlsx_parser(inputFile):
    ws = xlsx_opener(inputFile, 'Sheet1')
    contacts = load_data(ws)
    print(f'[+] Load {len(contacts)} from {inputFile}')

    return contacts


def csv_parser(inputFile):
    contacts = list()
    with open(inputFile, encoding='utf-8') as inF:
        print(f'[+] Open {inputFile}')
        data = csv.reader(inF, delimiter=';')
        header = next(data)
        for row in data:
            name, surname, phone1, phone2, workEmail, personalEmail, company, position = row
            contacts.append(
                {
                    'name': name,
                    'surname': surname,
                    'phone1': phone1,
                    'phone2': phone2,
                    'workEmail': workEmail,
                    'personalEmail': personalEmail,
                    'company': company,
                    'title': title
                }
            )

    print(f'[+] Load {len(contacts)} from {inputFile}')
    return contacts


def create_vcard(contact):
    template = f'''BEGIN:VCARD
VERSION:3.0
PRODID:-//Apple Inc.//Mac OS X 10.13.3//EN
N:{contact['surname']};{contact['name']};;;
FN:{contact['name']} {contact['surname']}
ORG:{contact['company']};
TITLE:{contact['title']}
EMAIL;type=INTERNET;type=WORK;type=pref:{contact['workEmail']}
EMAIL;type=INTERNET;type=HOME:{contact['personalEmail']}
TEL;type=CELL;type=VOICE;type=pref:{contact['phone1']}
TEL;type=CELL;type=VOICE:{contact['phone2']}
UID:e4aeda3b-aacf-42ea-8a2c-f35e7a0798f6
X-ABUID:97AF7201-206D-47B7-973F-BEA9FB7F70C5:ABPerson
END:VCARD'''

    return template


def csv_to_vcard(inputFile, outputFile):
    contacts = csv_parser(inputFile)
    with open(outputFile, 'w', encoding='utf-8') as outF:
        for contact in contacts:
            vcard = create_vcard(contact)
            outF.write(vcard + '\n')
    print(f'[+] Write {len(contacts)} to {outputFile}')


def xlsx_to_vcard(inputFile, outputFile):
    contacts = xlsx_parser(inputFile)
    with open(outputFile, 'w', encoding='utf-8') as outF:
        for contact in contacts:
            vcard = create_vcard(contact)
            outF.write(vcard + '\n')
    print(f'[+] Write {len(contacts)} to {outputFile}')


def converter(inputFile, outputFile):
    fileExtension = os.path.splitext(inputFile)[1]
    if fileExtension == '.csv':
        csv_to_vcard(inputFile, outputFile)
    elif fileExtension == '.xlsx':
        xlsx_to_vcard(inputFile, outputFile)
    else:
        print(f'[-] Unsupported file extension: {inputFile}')


def main():
    parser = argparse.ArgumentParser(description='Create Vcard from CSV or XLSX files.')
    parser.add_argument('-f', dest='inputFile', action='store', help='Input file contacts')
    parser.add_argument('-o', dest='outputFile', action='store', help='Output vcard file')

    args = parser.parse_args()
    converter(args.inputFile, args.outputFile)


if __name__ == '__main__':
    main()