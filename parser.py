from openpyxl import load_workbook
from pybtex.database.output.bibtex import Writer
from pybtex.database.input import bibtex
from transliterate import translit
import re
from copy import copy
from langdetect import detect, DetectorFactory
DetectorFactory.seed = 0

# The default for this library is to use quotation marks instead of curly braces
class CustomBibtexWriter(Writer):
    def quote(self, s):
        return '{' + s + '}'

# Method for styling hyperlink cells when adding them
def hyperlinkStyling():
    # set the last cell to be a hyperlink
    last_row = sheet.max_row
    cell = sheet.cell(row=last_row, column=5)
    left_cell = sheet.cell(row=last_row, column=4)

    # styling for hyperlink
    new_font = copy(left_cell.font)
    new_font.underline = "single"
    new_font.color = "0000FF"  # blue color in hex
    cell.font = new_font
    cell.hyperlink = link

bibtex_file = 'bells-2024-16.bib'
excel_file = "ORCID-FIL.xlsx"

parser = bibtex.Parser()
bib_data = parser.parse_file(bibtex_file)

bibtex_dict = {}
for entry_key, entry in bib_data.entries.items():
    orcid_pattern = r"\b(\d{4}-\d{4}-\d{4}-\d{3}[\dX])\b" # find the individual orcids in the entry
    orcids = re.findall(orcid_pattern, entry.fields['orcid'])
    for i, author in enumerate(entry.persons['author']):
        first = " ".join(author.first_names) if author.first_names else ''
        last = " ".join(author.last_names) if author.last_names else ''
        first = first.replace("Ð", "Đ").replace("ð", "đ") # Ð and Đ are of different unicode values...
                                                                                   # Poludeo sam tražijući problem haha
        last = last.replace("Ð", "Đ").replace("ð", "đ")
        if len(orcids) > i :
            bibtex_dict[(first, last)] = orcids[i]
        else:
            bibtex_dict[(first, last)] = None

wb = load_workbook(excel_file)
sheet = wb.active

# Remove empty rows from the sheet
for row in range(sheet.max_row, 0, -1):
    if all(cell.value in (None, "") for cell in sheet[row]):
        sheet.delete_rows(row)

headers = [cell.value for cell in sheet[1]]  # Header row
headers = [s.strip() if s is not None else None for s in headers] #strip header whitespace

# indexes of the rows
firstname_index = headers.index("Име")
lastname_index = headers.index("Презиме")
orcid_index = headers.index("ORCId идентификатор")

excel_dict = {}
i = 2
# Iterate over each row
for row in sheet.iter_rows(min_row=2, values_only=True):
    first_cyr = row[firstname_index]  # First name in Cyrillic
    last_cyr = row[lastname_index]  # Last name in Cyrillic
    orcid = row[orcid_index]  # orcid
    if first_cyr is None or last_cyr is None:
        continue
    # turn cyrillic to latin
    first_lat = translit(first_cyr, 'sr', reversed=True)
    last_lat = translit(last_cyr, 'sr', reversed=True)

    # If the bibtex file had the orcid for this person, but the excel doesn't, add it
    if orcid is None and bibtex_dict.keys().__contains__((first_lat, last_lat)) and bibtex_dict[(first_lat, last_lat)] is not None:
        sheet.cell(row=i, column=orcid_index + 1).value = bibtex_dict[(first_lat, last_lat)]
        link = ""
        if orcid is not None:
            link = f"https://orcid.org/{orcid}"
        hyperlinkStyling()
        print("Modified " + first_lat + " " + last_lat + " from Bibtex to excel")
    excel_dict[(first_lat, last_lat)] = orcid
    i += 1

common_names = set(bibtex_dict.keys()) & set(excel_dict.keys()) # List of names in both the bibtex and Excel
modify_bibtex = {}
for name in common_names:
    if bibtex_dict[name] is None and excel_dict[name] is not None: # There is data in the excel, but not bibtex
        modify_bibtex[name] = excel_dict[name]

# add excel data to bibtex
for entry_key, entry in bib_data.entries.items():
    orcid_pattern = r"\b(\d{4}-\d{4}-\d{4}-\d{3}[\dX])\b" # find the individual orcids in the entry
    for i, author in enumerate(entry.persons['author']):
        first = " ".join(author.first_names) if author.first_names else ''
        last = " ".join(author.last_names) if author.last_names else ''
        if modify_bibtex.__contains__((first, last)):
            if len(re.findall(orcid_pattern, entry.fields['orcid'])) > 0:
                entry.fields['orcid'] = entry.fields['orcid'] + " and " + modify_bibtex[(first, last)]
            else:
                entry.fields['orcid'] = modify_bibtex[(first, last)]
            print("modified " + first + " " + last + " in the bibtex file")


# People in the bibtex, not in the Excel
add_to_excel = set(bibtex_dict.keys()) - set(excel_dict.keys())
for name in add_to_excel:
    # Transliterating into Cyrillic causes issues, here is somewhat of a solution
    if detect(name[0] + " " + name[1]) in ['hr', 'sl']:  # if this classifies names as Croatian or Slovenian, transcribe
        first = translit(name[0], 'sr', reversed=False)
        last = translit(name[1], 'sr', reversed=False)
    else:
        first = name[0]
        last = name[1]
        # print(name[0] + " " + name[1] + " or " + first_cyr + " " + last_cyr + " is " + detect(name[0] + " " + name[1])) # testing

    orcid = bibtex_dict[name]
    link = ""
    if orcid is not None:
        link = f"https://orcid.org/{orcid}"
    sheet.append(["", first, last, orcid, link])
    hyperlinkStyling()

# Save excel changes a new file (can be changed to the same file to just edit rather than make a new one)
destination_file = "test.xlsx"
wb.save(destination_file)

# Save bibtex changes to a new file (can be changed to the same file to just edit rather than make a new one)
writer = CustomBibtexWriter()
modified_bibtex_file = "test.bib"
writer.write_file(bib_data, modified_bibtex_file)
