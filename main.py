from dotenv import load_dotenv
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import os
import argparse
import requests
import urllib.parse
import time

# https://positionstack.com/documentation

api_key = None

def get_coordinates(address:str, zipcode:str, city:str):

    url = "http://api.positionstack.com/v1/forward"
    url += f"?access_key={api_key}"
    q = f"{address}, {zipcode} {city}"
    url += f"&query={urllib.parse.quote(q)}"
    response = requests.get(url)
    if response.status_code == 200:
        js = response.json()
        if "data" in js:
            data = js["data"]
            if len(data) >= 1:
                rec1 = data[0]
                if rec1 == []:
                    return None
                return rec1
    else:
        print(f"Error {response.status_code} : {url}")
    
    return None

def add_cell(sheet, column, row, value, bold = False):
    char = get_column_letter(column)
    c = sheet[char + str(row)]
    c.value = value
    if bold:
        c.font = Font(bold = True)

def process(infile:str, outfile:str):
    print("Reading...")
    workbook = load_workbook(infile)
    worksheet = workbook.active

    workbook_out = Workbook()
    worksheet_out = workbook_out.active
    worksheet_out.title = "GEO"

    add_cell(worksheet_out, 1, 1, "Name", bold = True)
    add_cell(worksheet_out, 2, 1, "Address", bold = True)
    add_cell(worksheet_out, 3, 1, "PostalCode", bold = True)
    add_cell(worksheet_out, 4, 1, "City", bold = True)
    add_cell(worksheet_out, 5, 1, "Phone", bold = True)
    add_cell(worksheet_out, 6, 1, "Email", bold = True)
    add_cell(worksheet_out, 7, 1, "Latitude", bold = True)
    add_cell(worksheet_out, 8, 1, "Longitude", bold = True)
    add_cell(worksheet_out, 9, 1, "Confidence", bold = True)
    add_cell(worksheet_out, 10, 1, "Label", bold = True)

    print("Processing...")
    line = 2
    #line = 789
    while True:
        name = worksheet[f"A{line}"].value
        if name == "" or name is None:
            break
        address = worksheet[f"B{line}"].value
        zipcode = worksheet[f"C{line}"].value
        city = worksheet[f"D{line}"].value
        phone = worksheet[f"E{line}"].value
        email = worksheet[f"F{line}"].value

        add_cell(worksheet_out, 1, line, name)
        add_cell(worksheet_out, 2, line, address)
        add_cell(worksheet_out, 3, line, zipcode)
        add_cell(worksheet_out, 4, line, city)
        add_cell(worksheet_out, 5, line, phone)
        add_cell(worksheet_out, 6, line, email)

        try:
            coords = get_coordinates(address, zipcode, city)
            if coords is not None:
                latitude = coords["latitude"]
                longitude = coords["longitude"]
                label = coords["label"]
                confidence = coords["confidence"]

                print(f"{name} => {latitude},{longitude}")

                add_cell(worksheet_out, 7, line, latitude)
                add_cell(worksheet_out, 8, line, longitude)
                add_cell(worksheet_out, 9, line, confidence)
                add_cell(worksheet_out, 10, line, label)
            else:
                print(f"{name} => ??????????????????????")
        except Exception as e:
            print(f"Error: Line {line} - {name}")
            print(e)
        else:
            pass
        finally:
            pass

        time.sleep(0.5)
        line += 1
        #break

    print("Writing...")
    workbook_out.save(outfile)

def main():
    global api_key
    load_dotenv()
    env_path = Path('.')
    env_file = '.env'
    load_dotenv(dotenv_path = os.path.join(env_path, env_file))
    api_key = os.getenv("API_KEY")

    parser = argparse.ArgumentParser()
    parser.add_argument("inputfile", help = "Name of the input Excel file")
    parser.add_argument("outputfile", help = "Name of Excel file to write results to")
    args = parser.parse_args()

    if os.path.isfile(args.inputfile) == False:
        print(f"Error: file not found: {args.inputfile}")
        return

    infile = args.inputfile
    outfile = args.outputfile

    if infile[-5:].lower() != ".xlsx":
        print("Error: inputfile must be an .xlsx file")
        return

    if outfile[-5:].lower() != ".xlsx":
        outfile += ".xlsx"

    process(infile, outfile)

if __name__ == "__main__":
    main()

