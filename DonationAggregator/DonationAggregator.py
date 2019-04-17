import openpyxl
import argparse

# TODO: Include support for csv files using the built in Python CSV parser
# TODO: Progress indicator for huge files?


def startup():

    args = produce_parser()

    try:
        workbook = openpyxl.load_workbook(args.filename)
    except openpyxl.utils.exceptions.InvalidFileException:
        print("This program accepts only .xlsx, .xlsm, .xltx, and .xltm\nTry a different file.")
        return
    except FileNotFoundError:
        print("This file could not be opened. Try a different file.")
        return
    except RuntimeError as error:
        print("An unknown error occurred", error)
        return

    # This first sheet will always exist, since a workbook cannot exist without a sheet
    aggregated_donations = analyze(workbook[workbook.sheetnames[0]], args.name, args.committee, args.donation)

    if aggregated_donations is None:
        print("Either this file contains no data, or something unexpected went wrong.")
        return

    save_result(args.filename, workbook, aggregated_donations)


def produce_parser():
    parser = argparse.ArgumentParser(description="This tool is designed to simplify the process of analyzing and "
                                                 "interpreting the data found in FEC donation information by "
                                                 "aggregating donations for each person listed in the given dataset.")
    parser.add_argument("filename", help="the file to read donation data from")
    parser.add_argument("-n", "--name", default="O", metavar="", help="the column where the donors name can be found")
    parser.add_argument("-c", "--committee", default="B", metavar="", help="the column where the committee "
                                                                           "receiving the donation can be found")
    parser.add_argument("-d", "--donation", default="AI", metavar="", help="the column where the amount donated "
                                                                           "can be found")
    return parser.parse_args()


def analyze( data_sheet, name_col, committee_col, donation_col):

    # Check the sheet format
    if data_sheet[name_col + "1"].value != "contrib_name_adj" or data_sheet[committee_col + "1"].value != "committee_name" or \
            data_sheet[donation_col + "1"].value != "contribution_receipt_amount":
        print("This file is improperly formatted. Check the file and try again.")
        return

    aggregated_donations = {}

    # Note: this does not account for typos and misspelled names
    for index in range(data_sheet.min_row + 1, data_sheet.max_row + 1):

        name, org, don_amt = parse_row(data_sheet[index], name_col, committee_col, donation_col)

        # TODO: This is my entire data structure here: worth optimizing once it works
        # Ignore all rows that have either no org, no name, or no donation
        if name and org and don_amt:
            if name not in aggregated_donations:
                aggregated_donations[name] = {}

            if org not in aggregated_donations[name]:
                aggregated_donations[name][org] = 0

            aggregated_donations[name][org] += don_amt

    return aggregated_donations


def parse_row(row, name_col, committee_col, donation_col):
    # (name, organization, donation amount)
    # The indexes have to be done like this because openpyxl returns a tuple of columns
    return row[letter_number(name_col)].value, row[letter_number(committee_col)].value, \
           row[letter_number(donation_col)].value


# TODO: There is certainly a way to do this with more algorithmic ~sizzle~
def letter_number(col):

    if len(col) == 1:
        return ord(col) % 65
    elif len(col) == 2:
        return (((ord(col[0]) % 65) + 1) * 26) + (ord(col[1]) % 65)
    else:
        return None


def save_result(filename, workbook, aggregated_donations):

    result_sheet = workbook.create_sheet("aggregate_data")

    result_sheet["A1"] = "donor_name"
    result_sheet["B1"] = "committee_name"
    result_sheet["C1"] = "aggregate_amount"
    result_sheet["D1"] = "department"

    departmentList = {}
    directorySheet = workbook["DirectoryResults"]

    # Build the reference set of departments using the information in DirectoryResults
    for index in range(directorySheet.min_row + 1, directorySheet.max_row + 1):
        if directorySheet["B{0}".format(index)].value is not None:
            departmentList[directorySheet["A{0}".format(index)].value] = directorySheet["B{0}".format(index)].value

    # Start at 2, to account for the Excel double offset (start at 1 and a header row)
    curr_row = 2

    for (name, donation_entry) in aggregated_donations.items():
        if name in departmentList:
            for org in donation_entry:

                # Output goes to rows A, B, C, D, E
                result_sheet["A{}".format(curr_row)] = name
                result_sheet["B{}".format(curr_row)] = org
                result_sheet["C{}".format(curr_row)] = donation_entry[org]
                result_sheet["D{}".format(curr_row)] = departmentList[name]
                curr_row += 1

    print("Saving result...")
    workbook.save(filename)
    return


if __name__ == '__main__':
    startup()
