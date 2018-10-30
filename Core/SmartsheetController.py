import smartsheet
import logging
import collections as co

bug_link_prefix = "http://bugzilla.lln.iba/show_bug.cgi?id="
access_token = "vwlyghxsunr2uzmotowt4wr7ei"

# read the map file into memory
MAP = list()
with open("../map.txt") as f:
    for line in f.readlines():
        mid, bug = line.split("--")
        MAP.append(mid)
        MAP.append(str(bug).strip('\n'))

print MAP

sheet_map = {}

# Initialize client
ss = smartsheet.Smartsheet(access_token)
# Make sure we don't miss any error
ss.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename='smartsheet_req.log', level=logging.INFO)

action = ss.Sheets.list_sheets(include_all=True)

SHEET_ID = ""
for sheet in action.data:
    sheet_map[str(sheet.name)] = str(sheet.id)
    SHEET_ID = sheet_map.get("02 - Guangzhou - Test Matrix")
print "got the sheet 'Guangzhou - Test Matrix' id = " + str(sheet_map.get("02 - Guangzhou - Test Matrix"))

# get columns
columns = ss.Sheets.get_columns(SHEET_ID)
mid_column = ss.Sheets.get_column_by_title(SHEET_ID, "MID#")
bug_column = ss.Sheets.get_column_by_title(SHEET_ID, "BUG#")
print "got mid column= %s & bug column= %s" % (mid_column.id, bug_column.id)

# get all mid cells and feedback all the mid cells row number
mys = ss.Sheets.get_sheet(SHEET_ID, column_ids=[mid_column.id])
NB = 0
MID_row_set = co.OrderedDict()
for myr in mys.rows:
    if myr.row_number >= 3:
        for myc in myr.cells:
            if myc.display_value is not None and myc.display_value != u"72997":
                # find all MID# row_id
                NB += 1
                MID_row_set[myr.id] = myc.display_value

print "totle number of MID# is: ", NB


def find_bug_number(MID_nb):
    for i in range(len(MAP)-1):
        if i % 2 == 1 and MAP[i+1] == MID_nb:
            MAP.pop(i)
            return MAP.pop(i+1)

def find_mid_number(row_id):
    return MID_row_set.get(row_id)


def build_cell_content(row_id, bug_number):
    # build new cell value
    newCell = ss.models.Cell()
    newCell.column_id = bug_column.id
    newCell.value = bug_number
    newHyperlink = ss.models.Hyperlink()
    newHyperlink.url = bug_link_prefix + bug_number
    newCell.hyperlink = newHyperlink

    # build the row to update
    newRow = ss.models.Row()
    newRow.id = row_id
    newRow.cells.append(newCell)
    return newRow


def update_multi_rows(rows):
    result = ss.Sheets.update_rows(SHEET_ID, rows)



BUG_row_id_set = []
rows_to_update = []
# iterate the BUG# column
print "Start to write the BUG number into Smartsheet"
bugs = ss.Sheets.get_sheet(SHEET_ID, column_ids=[bug_column.id])
for bugr in bugs.rows:
    if bugr.row_number >= 3:
        for bugc in bugr.cells:
            if bugc.value is None:
                row_id = bugr.id
                # find the MID number by row_id
                mid_nb = find_mid_number(row_id)
                # find the BUG number from map file
                if mid_nb is not None:
                    bug_nb = find_bug_number(mid_nb)
                    print "Start to looking for MID nb = ", mid_nb, "BUG nb = ", bug_nb
                    rows_to_update.append(build_cell_content(row_id, bug_nb))
# update
update_multi_rows(rows_to_update)
