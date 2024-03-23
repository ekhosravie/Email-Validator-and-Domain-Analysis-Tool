Email Validator and Domain Analysis Tool
This project consists of a Python script and a Google Apps Script for validating and analyzing email addresses from an Excel file, leveraging the Mailboxlayer API and Google Sheets for advanced interaction.

Table of Contents
Overview
Python Script
Google Apps Script
Usage
Contributing
License
Overview

The project aims to provide a comprehensive solution for validating and analyzing email addresses from an Excel file. It consists of two main components:

Python Script (email_validator.py): This script reads an Excel file containing email addresses, validates them using the Mailboxlayer API, extracts email domains, and performs keyword matching. It then uploads the analyzed data to Google Sheets.

Google Apps Script (email_analysis.gs): This script is integrated with Google Sheets and provides functionalities for running the analysis, formatting the sheet, and adding a pie chart for visualizing the analysis results.

Python Script
The email_validator.py script performs the following tasks:

Loads an Excel file containing email addresses and keywords for matching.
Validates emails via the Mailboxlayer API.
Extracts email domains and performs keyword matching.
Uploads the analyzed data to Google Sheets.
Google Apps Script
The email_analysis.gs script provides the following functionalities:

Custom menu options for running analysis, formatting the sheet, and adding a pie chart.
Analysis function (runAnalysis) for processing email addresses.
Formatting function (formatSheet) for applying styles to the sheet based on email validation results.
Pie chart function (addPieChart) for visualizing the percentage of valid and invalid email addresses.
Usage
To use this tool, follow these steps:

Clone the repository to your local machine.
Install the necessary Python dependencies (Pandas, Requests, Google API).
Replace placeholder API keys and file paths with your actual credentials and file paths.
Run the Python script to perform email validation and analysis.
Open the Google Sheets linked to your Google account.
Access the custom menu options under "Email Analysis" to run the analysis, format the sheet, and add a pie chart.
