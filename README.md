# pyExcel Analysis

This project analyzes aircraft and animal collision data from an Excel spreadsheet using Python's `openpyxl` library. The analysis results are saved in a new Excel workbook with various charts.

## Overview

The script processes data from the `aircraftWildlifeStrikes.xlsx` file and performs the following analyses:
- Identifies the animal involved in the most collisions.
- Determines the year with the most collisions.
- Finds the month with the most collisions.
- Identifies the airline company with the most animal collisions.

The results are saved in a new workbook, `aircraftWildlifeAnalysis.xlsx`, with a separate sheet and chart for each analysis.

## Features

- **Data Cleaning**: Standardizes animal names and handles missing or unknown values.
- **Data Counting**: Counts occurrences of each category (animal, year, month, airline).
- **Chart Generation**: Creates bar charts to visualize the data.
- **Excel Integration**: Uses `openpyxl` to manipulate Excel files and create charts.

## Usage

1. Ensure you have the required dependencies installed:
   ```sh
   pip install openpyxl
   ```
