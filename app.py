
from datetime import date
from pandas import ExcelWriter

import sys

import pandas as pd

data = {}

def getProductId():
    productId = input("Please enter the product id: ")
    sod = data["Sales Order Detail"]
    try:
        val = int(productId)
        if(val in sod["ProductID"]):
            updateUnitPrice(val)
        else:
            getProductId()
    except ValueError:
        getProductId()


def updateUnitPrice(productID):
    sod = data["Sales Order Detail"]
    sod.loc[sod.ProductID == productID, 'UnitPrice'] = sod["UnitPrice"] / 2
    print("Successfully updated unit price for product " + str(productID))
    cont()


def deDuplicate():
    data["Sales Order Header"].drop_duplicates(["SalesOrderID"], keep="first", inplace=True)
    cont()

def isNaN(num):
    return num == num


def roundAmounts():
    sheetList = ["Sales Order Detail", "Sales Order Header"]

    colList = [["UnitPrice", "LineTotal"], ["SubTotal", "TaxAmt", "Freight", "TotalDue"]]

    for sheet in sheetList:
        df = data[sheet]
        index = sheetList.index(sheet)
        for col in colList[index]:
            df[col] = round(df[col], 2)

    print("Successfully rounded all dollar amounts")
    cont()

def checkInput(value):
    try:
        val = int(value)
        if(val in range(1, 4)):
            if(val == 1):
                getProductId()
            if(val == 2):
                deDuplicate()
            if(val == 3):
                roundAmounts()
        else:
            checkInput(input("Invalid input. Try again.: "))
    except ValueError:
        checkInput(input("Invalid input. Try again.: "))

def app():
    print("Which process do you want to run?")
    print("1.) Update Unit Price")
    print("2.) DeDuplicate Sales Order Header")
    print("3.) Round Dollar amounts")
    checkInput(input("(Use number as your input): "))

def cont():
    process = input("Would you like to do another process? (YES or NO): ")
    if(process.upper() == "YES"):
        app()
    else:
        if(process.upper() == "NO"):
            print("Your excel file is being saved. Please wait...")
            saveExcel()
        else:
            cont()

def saveExcel():
    today = date.today()
    d = today.strftime("%m-%d-%Y")
    writer = ExcelWriter('Jacob Bartoletta – FIT Sales Data Date('+str(d)+').xlsx')
    for key in data:
        data[key].to_excel(writer, key, index=False)

    writer.save()
    print("Your excel file has successfully been saved.")

noParams = len(sys.argv) - 1
if(noParams == 1):
    excel_worksheet = sys.argv[1]
    print("Data loading please wait...")
    data = pd.read_excel(excel_worksheet, sheet_name=None)
    app()
else:
    print("Invalid arguments.")

