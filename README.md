# Pole Mews Project

This project involves data analysis and management for sales forecasting and inventory optimization using Python. The main objective is to analyze historical sales data, forecast future sales, calculate inventory metrics, and classify products based on their sales performance.

## Overview

The project involves the following key steps:

1. Importing libraries and setting up the environment.
2. Importing data from Google Sheets and local Excel files.
3. Data preprocessing and cleaning.
4. Analysis of sales data and inventory.
5. Calculation of inventory metrics such as standard deviation, stock roulant, and ABC classification.
6. Updating the Google Sheets file with the analyzed data.

## Files

- `pole_projet_mews.ipynb`: Jupyter Notebook containing the Python code for the project.
- `DataModeleComplet.xlsx`: Excel file containing the raw data used in the analysis.
- `README.md`: This file, providing an overview of the project.

## Usage

To use this project:

1. Ensure you have the required Python libraries installed. You can find the list of dependencies in the Jupyter Notebook.
2. Upload the `DataModeleComplet.xlsx` file to the appropriate directory (`Mews_Partners`) in your working environment.
3. Run the Jupyter Notebook (`pole_projet_mews.ipynb`) cell by cell to execute the code and perform the analysis.
4. After the analysis is complete, the updated data will be exported to the Google Sheets file named `Pole_Mews`.

## Dependencies

- pandas
- matplotlib
- numpy
- unidecode
- scipy
- openpyxl
- gspread
- google-auth
- oauth2client

These libraries can be installed using pip:
pip install pandas matplotlib numpy unidecode scipy openpyxl gspread google-auth oauth2client
