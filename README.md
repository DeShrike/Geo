# Forward Geo Location

Uses PositionStack API

## Installation

Clone this repo

Then run:

```
python3 -m venv venv
source venv/bin/activate (Linux)
venv\bin\activate.bat    (Windows)
pip3 install -r requirements.txt 
```

## Getting Started

Create an account at https://positionstack.com/

Rename sample_dot_env to .env 

Open .env in a text editor and add your API key.

Addresses are read from an Excel file and coordinates are written back to the same Excel file.

See sample_addresses.xlsx (the 'before') and sample_result.xlsx (the 'after').

Run:

```
python3 main.py <name of Excel file>
```

Options:

-u : Update the rows where longitude or latitude are empty. (the default)

-r : Update the rows where longitude or latitude are empty, or where the confidence is not equal to 1.

## Excel file format

Only the first worksheet is read.

Row 1 should be a header row.

- Column A: A name, not used during the geolocation
- Column B: The address
- Column C: The zip code
- Column D: The City
- Column E: Not used
- Column F: Not used
- Column G: Not used
- Column H: The country code

The results are written back to columns I-U:

- Column I: The longitude
- Column J: The latitude
- Column K: A confidence score
- Column L: Region
- Column M: Region Code
- Column N: County
- Column O: Locality
- Column P: Administrative Area
- Column Q: Neighbourhood
- Column R: Country
- Column S: Country Code
- Column T: Continent
- Column U: Label

