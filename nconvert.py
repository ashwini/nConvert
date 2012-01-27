#!/usr/bin/env python
#
# nConvert v1.1
# Author: ashwini@majestik.net
#
#
# This utility will convert a simple, human-readable Contacts CSV to a format that's
# compatible with Nokia PC Suite. Had to write a utility to do this because standard
# Microsoft Outlook CSV exports weren't reading in Nokia PC Suite. Figured out it was because
# NPCS wants all these extraneous fields and utf-16 encoding. Obviously. :)
#
# Format your Contacts CSV like this: "First Name","Last Name","Mobile Phone"
#
# Just run 'python nconvert.py' to start.
#

__author__ = 'ashwini@majestik.net'


print "nConvert v1.1 - Utility to convert CSV to a Nokia-CSV format for import in Nokia PC Suite\r\n"

inputfile = raw_input('Please enter the name of your CSV file (e.g. my_contacts.csv): ')

try:
    contacts = open(inputfile, 'rU')

except IOError as e:
    print "Unable to open \"{0}\" - {1})".format(inputfile, e)

else:
    filter = contacts.read()
    filter = filter.split("\n")
    contacts.close()


    # Remove header from source if exists
    for hfield in ["First name", "name", "Name"]:
        if hfield in filter[0]: del filter[0]

    # Figure out name of output file
    if "\\" in inputfile:
        filename = inputfile.split("\\")[-1]
        outputfile = inputfile.replace(filename, "Nokia-"+filename)
    elif "/" in inputfile:
        filename = inputfile.split("/")[-1]
        outputfile = inputfile.replace(filename, "Nokia-"+filename)
    else:
        filename = inputfile
        outputfile = "Nokia-"+inputfile

    # Create output file, write the Nokia header, encode in UTF-16
    contacts_out = open(outputfile, 'wb')
    header = "\"Title\",\"First name\",\"Middle name\",\"Last name\",\"Suffix\",\"Job title\",\"Company\",\"Birthday\",\"SIP address\",\"Push-to-talk\",\"Share view\",\"User ID\",\"Notes\",\"General mobile\",\"General phone\",\"General email\",\"General fax\",\"General video call\",\"General web address\",\"General VOIP address\",\"General P.O.Box\",\"General extension\",\"General street\",\"General postal/ZIP code\",\"General city\",\"General state/province\",\"General country/region\",\"Home mobile\",\"Home phone\",\"Home email\",\"Home fax\",\"Home video call\",\"Home web address\",\"Home VOIP address\",\"Home P.O.Box\",\"Home extension\",\"Home street\",\"Home postal/ZIP code\",\"Home city\",\"Home state/province\",\"Home country/region\",\"Business mobile\",\"Business phone\",\"Business email\",\"Business fax\",\"Business video call\",\"Business web address\",\"Business VOIP address\",\"Business P.O.Box\",\"Business extension\",\"Business street\",\"Business postal/ZIP code\",\"Business city\",\"Business state/province\",\"Business country/region\",\"\"{0}".format("\r\n")
    header = header.encode('utf-16')
    contacts_out.write(header)
    ContactCount = 0

    # Convert all contacts from source CSV to Nokia friendly output CSV
    for row in filter:
        try:
            field = row.split(",")

            # Clean up the phone number field
            for symbol in ["-", " ", "."]:
                if symbol in field[2]: field[2] = field[2].replace(symbol, "")

            new_row = "\"\",{0},\"\",{1},\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",{2},\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\"{3}".format(field[0], field[1], field[2],"\r\n")

            new_row = new_row.encode('utf-16')
            ContactCount += 1
            contacts_out.write(new_row)
        except IndexError:
            contacts_out.close()
            break

    print "\r\nSuccessfully created '{0}' with {1} total contacts. Go ahead and import that!".format(outputfile, ContactCount)

