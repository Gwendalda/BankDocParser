#!/usr/bin/Python3.9
import json
import sys

import xlsxwriter
import os

COLUMNS = {"DATE DE L'OPÉRATION": 0,
           "DATE D'INSCRIPTION": 1,
           "DESCRIPTION": 2,
           "MONTANT": 3,
           "PST": 4,
           "GST": 5,
           "HST": 6,
           "TOTAL TAXES": 7,
           "MONTANT AVANT TAXES": 8,
           "REGIME DE TAXE": 9,
           "COMPTE": 10
           }
base_path = getattr(sys, '_MEIPASS', os.path.dirname(
    os.path.abspath(__file__)))

# Construct the full path to the JSON file
ACC_HIST_PATH = os.path.join(base_path, 'accountHistory.json')


DATE_OPERATION_KEY = "date_operation"
DATE_INSCRIPTION_KEY = "date_inscription"
DESCRIPTION_KEY = "description"
MONTANT_KEY = "montant"
# must be a dictionary containing the 'num_format' key.
CURRENCY_FORMAT = {'num_format': '[$$-409]#,##0.00'}
DEFAULT_TAX_REGIME = "QC"
TAXE_RATES = {
    "QC": {
        "PST": 0.09975,
        "GST": 0.05,
        "HST": 0.00,
    },
    "AB": {
        "PST": 0.00,
        "GST": 0.05,
        "HST": 0.00,
    },
    "BC": {
        "PST": 0.07,
        "GST": 0.05,
        "HST": 0.00
    },
    "MA": {
        "PST": 0.07,
        "GST": 0.05,
        "HST": 0.00
    },
    "NB": {
        "PST": 0.00,
        "GST": 0.00,
        "HST": 0.15
    },
    "NL": {
        "PST": 0.00,
        "GST": 0.00,
        "HST": 0.15
    },
    "NT": {
        "PST": 0.00,
        "GST": 0.05,
        "HST": 0.00
    },
    "NS": {
        "PST": 0.00,
        "GST": 0.00,
        "HST": 0.15
    },
    "NA": {  # nunavut
        "PST": 0.00,
        "GST": 0.05,
        "HST": 0.00
    },
    "ON": {
        "PST": 0.00,
        "GST": 0.00,
        "HST": 0.13
    },
    "PEI": {
        "PST": 0.00,
        "GST": 0.00,
        "HST": 0.15
    },
    "SK": {
        "PST": 0.06,
        "GST": 0.05,
        "HST": 0.00
    },
    "YK": {
        "PST": 0.00,
        "GST": 0.05,
        "HST": 0.00
    },
    "None": {
        "PST": 0.00,
        "GST": 0.00,
        "HST": 0.00
    }

}


def jsonParser(jsonFilePath):  # file path has to be correct !
    with open(ACC_HIST_PATH, "r") as f:
        accountHistory = json.load(f)
    with open(jsonFilePath, "r") as f:
        jsondoc = json.load(f)
    data = jsondoc[0]["key_1"]
    writeExcelFile(jsonFilePath, data, accountHistory)


def writeDataInWorksheet(worksheet, data, currencyformat):
    # writing the first four columns of information into the file.
    for i in range(len(data)):
        worksheet.write(
            i + 1, COLUMNS["DATE DE L'OPÉRATION"], data[i][DATE_OPERATION_KEY])
        worksheet.write(
            i + 1, COLUMNS["DATE D'INSCRIPTION"], data[i][DATE_INSCRIPTION_KEY])
        worksheet.write_string(
            i + 1, COLUMNS["DESCRIPTION"], data[i][DESCRIPTION_KEY])
        worksheet.write_number(
            i + 1, COLUMNS["MONTANT"], float(data[i][MONTANT_KEY]), currencyformat)


def writeColumnNames(worksheet):
    for i in range(len(COLUMNS.keys())):
        # they key name for every column
        worksheet.write(0, i, list(COLUMNS.keys())[i])


def writeTotals(worksheet, currencyformat, rows):
    worksheet.write(rows + 2, 0, "TOTAUX")
    worksheet.write_formula(
        rows + 2, 3, "=SUM(D2:D{0})".format(rows + 1), currencyformat)
    worksheet.write_formula(
        rows + 2, 4, "=SUM(E2:E{0})".format(rows + 1), currencyformat)
    worksheet.write_formula(
        rows + 2, 5, "=SUM(F2:F{0})".format(rows + 1), currencyformat)
    worksheet.write_formula(
        rows + 2, 6, "=SUM(G2:G{0})".format(rows + 1), currencyformat)
    worksheet.write_formula(
        rows + 2, 7, "=SUM(H2:H{0})".format(rows + 1), currencyformat)
    worksheet.write_formula(
        rows + 2, 8, "=SUM(I2:I{0})".format(rows + 1), currencyformat)


def writeFromAccountHistory(worksheet, accountHistory, data, rows):
    for i in range(0, rows):
        try:
            worksheet.write(
                i+1, COLUMNS["COMPTE"], accountHistory[data[i]["description"]]["COMPTE"])
            worksheet.write(
                i+1, COLUMNS["REGIME DE TAXE"], accountHistory[data[i]["description"]]["TAX REGIME"])
        except Exception as e:
            pass


def writeDefaultTaxRegime(worksheet, rows):
    for i in range(1, rows):
        worksheet.write(i, COLUMNS["REGIME DE TAXE"], DEFAULT_TAX_REGIME)


def writeTaxRegimeDataValidation(worksheet, rows):
    worksheet.data_validation(1, COLUMNS["REGIME DE TAXE"], rows, COLUMNS["REGIME DE TAXE"], {"validate": "list",
                                                                                              "source": "=TaxInfo!$A$2:$A$15",
                                                                                              "ignore_blank": False})


def writeTaxCalculations(worksheet, currencyformat, rows):
    for i in range(rows - 1):
        formula = "=D{0}*INDEX(TaxInfo!B2:E16,MATCH(J{0},TaxInfo!A2:A16, 0),1)/(1+INDEX(TaxInfo!B2:E16,MATCH(J{0},TaxInfo!A2:A16, 0),4))".format(i + 2)
        worksheet.write_formula(i+1, COLUMNS["PST"], formula, currencyformat)
        formula = "=D{0}*INDEX(TaxInfo!B2:E16,MATCH(J{0},TaxInfo!A2:A16, 0),2)/(1+INDEX(TaxInfo!B2:E16,MATCH(J{0},TaxInfo!A2:A16, 0),4))".format(i + 2)
        worksheet.write_formula(i + 1, COLUMNS["GST"], formula, currencyformat)
        formula = "=D{0}*INDEX(TaxInfo!B2:E16,MATCH(J{0},TaxInfo!A2:A16, 0),3)/(1+INDEX(TaxInfo!B2:E16,MATCH(J{0},TaxInfo!A2:A16, 0),4))".format(i + 2)
        worksheet.write_formula(i + 1, COLUMNS["HST"], formula, currencyformat)
        formula = "=E{0}+F{0}+G{0}".format(i+2)
        worksheet.write_formula(
            i + 1, COLUMNS["TOTAL TAXES"], formula, currencyformat)
        formula = "=D{0}-H{0}".format(i+2)
        worksheet.write_formula(
            i + 1, COLUMNS["MONTANT AVANT TAXES"], formula, currencyformat)


def writeTaxTable(worksheet):
    writeTaxTableColumnNames(worksheet)
    for i in range(1, len(TAXE_RATES.keys()) + 1):
        province = list(TAXE_RATES.keys())[i-1]
        worksheet.write(i, 0, province)
        worksheet.write_number(i, 1, float(TAXE_RATES[province]["PST"]))
        worksheet.write_number(i, 2, float(TAXE_RATES[province]["GST"]))
        worksheet.write_number(i, 3, float(TAXE_RATES[province]["HST"]))
        worksheet.write_formula(i, 4, "=SUM(B{0}, C{0}, D{0})".format(i+1))


def writeTaxTableColumnNames(worksheet):
    worksheet.write(0, 0, "PROVINCE")
    worksheet.write(0, 1, "PST")
    worksheet.write(0, 2, "GST")
    worksheet.write(0, 3, "HST")
    worksheet.write(0, 4, "Total")


def writeAccountTotals(worksheet, rows, currencyFormat):
    worksheet.write(0, 0, "COMPTE")
    worksheet.write(0, 1, "TOTAL")
    worksheet.write(1, 0, "PST")
    worksheet.write(2, 0, "GST")
    worksheet.write(3, 0, "HST")
    formula = "=Depenses!E{0}".format(rows+3)
    worksheet.write_formula(1, 1, formula, currencyFormat)
    formula = "=Depenses!F{0}".format(rows+3)
    worksheet.write_formula(2, 1, formula, currencyFormat)
    formula = "=Depenses!G{0}".format(rows+3)
    worksheet.write_formula(3, 1, formula, currencyFormat)
    formula = '=UNIQUE(FILTER(Depenses!K2:K{0},Depenses!K2:K{0}<>""))'.format(
        rows)
    worksheet.write_formula(4, 0, formula)
    formula = '=SUMIF(Depenses!K2:K{0},UNIQUE(FILTER(Depenses!K2:K{0},Depenses!K2:K{0}<>"")),Depenses!I2:I{0})'.format(
        rows)
    worksheet.write_formula(4, 1, formula, currencyFormat)


def writeExcelFile(jsonFilePath, data, accountHistory):
    workbook = xlsxwriter.Workbook(
        jsonFilePath.replace(".json", ".xlsx"))  # creating the .xlsx file
    depensesWorksheet = workbook.add_worksheet(
        "Depenses")  # adding a worksheet to work on
    taxInfoWorksheet = workbook.add_worksheet("TaxInfo")
    accountTotals = workbook.add_worksheet("accountTotals")
    currencyformat = workbook.add_format(
        CURRENCY_FORMAT)  # defining the currency format
    writeColumnNames(depensesWorksheet)
    writeDataInWorksheet(depensesWorksheet, data, currencyformat)
    # adjusting the length of the rows to account for the column names
    rows = len(data) + 1
    writeTaxTable(taxInfoWorksheet)
    writeTaxRegimeDataValidation(depensesWorksheet, rows)
    writeDefaultTaxRegime(depensesWorksheet, rows)
    writeTaxCalculations(depensesWorksheet, currencyformat, rows)
    writeTotals(depensesWorksheet, currencyformat, rows)
    # overwrite from memory.
    writeFromAccountHistory(depensesWorksheet, accountHistory, data, rows)
    writeAccountTotals(accountTotals, rows, currencyformat)
    workbook.close()
