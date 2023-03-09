# Table Cards Generator

This program generates a PowerPoint presentation of table cards using a given Excel file containing the names and titles of the attendees, and a template PowerPoint file.

## Installation

This program requires `Python` and the `pptx` package. To install the necessary packages, run:

pip install pptx pandas


## Usage

To use this program, follow these steps:

1. Prepare an Excel file containing the names and titles of the attendees, with the first column containing the names and the second column containing the titles.

2. Prepare a PowerPoint template file with the desired layout for the table cards.

3. Run the `table_cards_generator.py` script, passing in the Excel file and the template PowerPoint file as arguments:

python table_cards_generator.py --excel_file example.xlsx --pptx_file template.pptx


4. The program will generate a new PowerPoint file with table cards for each attendee, based on the template file.

## Acknowledgements

This program was developed using the `pptx` package by scanny.
