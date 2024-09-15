# Automated Data Extraction and Transformation with Power Query

## Overview
This repository contains a Power Query script that automates the extraction, cleaning, and transformation of Excel files from a specified folder. The script performs the following key tasks:

- **Folder-level data extraction:** Extracts data from the latest modified Excel file in a given folder.
- **Sheet filtering and selection:** Automatically selects sheets based on specific naming conventions and filters out unnecessary data.
- **Dynamic transformations:** Renames and merges columns, adds new columns, and applies conditional filters to ensure data consistency.

## Features

- **File Selection:** Automatically selects the most recent file based on the "Date Modified" attribute.
- **Sheet Filtering:** Excludes files and sheets that start with specific characters like `_` or `~`.
- **Data Transformation:** Applies multiple transformations such as:
  - Renaming columns
  - Merging columns
  - Adding location and contract details dynamically
  - Filtering transmittals based on specific criteria
- **Advanced M Code:** Demonstrates the use of advanced Power Query M code for complex data manipulations, including:
  - `Table.TransformColumns`
  - `Table.SelectRows`
  - `Table.CombineColumns`
  