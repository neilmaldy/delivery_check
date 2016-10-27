import openpyxl
import os
import sys
import win32com.client

email_part_1 = """Greetings,

You are listed as delivery contact for site """

email_part_2 = """.  Can you please check the following and notify us of any corrections:

"""

email_part_3 = """

You may receive multiple emails if there are variations in the details for this site.  Please let us know the most accurate information so we can update our records.

Thank you,
Neil Maldonado
Support Account Manager
NetApp
"""


def send_mail_via_com(text, subject, recipient):
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = subject
    newMail.Body = text
    # newMail.HTMLBody  = htmltext
    newMail.To = recipient
    # newMail.CC = "mneil@netapp.com"
    # newMail.BCC = "mneil@netapp.com"
    # attachment1 = "c:\\mypic.jpg"
    # newMail.Attachments.Add(attachment1)
    newMail.Send()    

# send_mail_via_com("email text", "email subject", "mneil@netapp.com")


def email_delivery_contacts(logistics_file):

    # check for file
    if not os.path.isfile(logistics_file):
        print("Could not find quote file " + logistics_file, file=sys.stderr)
        return

    # read quote from quote.xlsx
    wb = openpyxl.load_workbook(logistics_file, read_only=True)
    sheet = wb.active
    rows = sheet.rows

    # column headings in header_row
    header_row = [cell.value for cell in next(rows)]

    while header_row[0] != "Serial Number Owner Name":
        header_row = [cell.value for cell in next(rows)]

    logistics_rows = []
    emails = set()
    sites = set()
    # put remaining rows in logistics_rows
    for row in rows:
        record = {}

        # store each row in record
        for key, cell in zip(header_row, row):

            if cell.value is None:
                record[key] = ''
            elif cell.data_type == 's':
                # strip extra spaces from strings
                record[key] = cell.value.strip()
            else:
                # store everything else (numbers and dates) unchanged
                record[key] = cell.value

        # add row/record to logistics_rows
        logistics_rows.append(record)
        email_text = email_part_1 + \
            record['Installed At Site Name'] + \
            email_part_2 + \
            record['Delivery Contact Name'] + '\n' + \
            record['Delivery Contact Phone'] + '\n' + \
            record['Delivery Contact eMail'] + '\n' + \
            record['Logistics Ship To Address Party Name 1'] + '\n' + \
            record['Logistics Address'] + '\n' + \
            record['Logisitcs City'] + '\n' + \
            record['Logistics State/Province'] + '\n' + \
            record['Logistic Postal Code'] + '\n' + \
            record['Logistics Country'] + '\n' + \
            "Receiving Hours: " + record['Goods Receiving Hour'] + '\n' + \
            email_part_3
        if "UNKNOWN" not in record['Delivery Contact eMail'].upper() and email_text.lower() not in emails:
            emails.add(email_text.lower())
            print(email_text, file=sys.stderr)
            if record['Installed At Site Name'] not in sites:
                sites.add(record['Installed At Site Name'])
            else:
                print("Found multiple delivery contacts for site " + record['Installed At Site Name'], file=sys.stderr)
    print("Done.", file=sys.stderr)
    return

if __name__ == "__main__":
    print("Attempting to process logistics.xlsx", file=sys.stderr)
    email_delivery_contacts("logistics.xlsx")
