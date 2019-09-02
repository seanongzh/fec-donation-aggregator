# fec-donation-aggregator

## Background
The Donation Aggregator is a Python tool to compile FEC donation data for easier analysis and interpretation. Developed for the Tufts Daily.

When donation information is pulled from the FEC, the information is pretty formidable. Rows and rows of data, each one detailing an individual donation by a single person.

To address this, I wrote the Donation Aggregator. This program parses the data downloaded from the FEC and sums together each person's total contributions to a given political committee. Additionally, the script will attempt to make a determination on what party the political committee is associated with, using the FEC's wonderful [publicly available API](https://api.open.fec.gov/developers/). As an added bonus, the current iteration of the program (specifically tuned for use by the Tufts Daily team) uses the Tufts directory to determine the department affiliation of each person in the list, unlocking the ability to compare political donations across campus.

## Usage

Example usage (from ``/DonationAggregator/``):

`python DonationAggregator.py ../spreadsheets/Mini_test.xlsx ../spreadsheets/real_files/DirectoryResults_2017-2018.xlsx`

Running the script in this way will cause the input file (Mini_test.xlsx) to be modified with a new sheet containing the aggregated information from the first sheet in the document.

For example, three separate donations by one person to the same committee will be combined into a single row in the new sheet, making for much simpler data analysis.

There are also a few scripts included that provide various secondary functionality that further supplements the data (with specific tools designed for Tufts University).

## Advanced Usage

TODO: Explain the usage of each of the additional scripts, and what steps need to be taken to fully prepare the data for analysis as the program expects.

DISCLAIMER: While personally identifying information is included in the `/spreadsheets/` folder, all of this information has already been made public by the FEC. No information that was not already publicly available has been released in this repository.
