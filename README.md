# Forward Geo Locating 

Uses PositionStack API

## Install

Clone this repo

Then run:

'''
python3 -m venv env
source env/bin/activate
pip3 install -r requirements.txt 
'''

## Getting Started

Create an account at https://positionstack.com/

Put your API key in .env (see sample_dot_env)

Addresses are read from an Excel file

Results are written to an Excel file

See sample_addresses.xlsx and sample_results.xlsx

Run:

'''
python3 main.py addressfile.xlsx outfile.xlsx
'''

