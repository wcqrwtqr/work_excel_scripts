from pathlib import Path  # Standard Python Module
import csv
from openpyxl import load_workbook, Workbook  # pip install openpyxl


SOURCE_DIR = "/Users/mohammedalbatati/Downloads/excel-script"
excel_files = list(Path(SOURCE_DIR).glob("*.xlsx"))
# dict_values = {}

for excel_file in excel_files:
    wb = load_workbook(filename=excel_file, read_only=True)

    supervisorName = wb.worksheets[0]["C6"].value
    date = wb.worksheets[0]["I5"].value
    client = wb.worksheets[0]["F5"].value
    well = wb.worksheets[0]["F7"].value
    jobID = wb.worksheets[0]["I7"].value
    FlowingHour = wb.worksheets[0]["N23"].value
    MaxOilRate = wb.worksheets[0]["N24"].value
    avgOilRate = wb.worksheets[0]["N25"].value
    MaxWaterRate = wb.worksheets[0]["N26"].value
    avgWaterRate = wb.worksheets[0]["N27"].value
    avgGasRate = wb.worksheets[0]["N28"].value
    staticPressure = wb.worksheets[0]["N29"].value
    diffPressure = wb.worksheets[0]["N30"].value
    CMF = wb.worksheets[0]["N31"].value
    H2S = wb.worksheets[0]["N33"].value
    CO2 = wb.worksheets[0]["N34"].value
    BSW = wb.worksheets[0]["N35"].value
    WHP = wb.worksheets[0]["N36"].value
    WHT = wb.worksheets[0]["N37"].value
    choke = wb.worksheets[0]["N38"].value


    dict_values = {
        "supervisor": supervisorName, "date": date, "client": client, "well": well, "jobID": jobID,
        "FlowingHour": FlowingHour, "MaxOilRate": MaxOilRate, "avgOilRate": avgOilRate,
        "MaxWaterRate": MaxWaterRate, "avgWaterRate": avgWaterRate, "avgGasRate": avgGasRate,
        "staticPressure": staticPressure, "diffPressure": diffPressure, "CMF": CMF, "H2S": H2S,
        "CO2": CO2, "BSW": BSW, "WHP": WHP, "WHT": WHT, "choke": choke, "fileName": excel_file
    }

    with open('db.csv', 'a') as csv_file:
        writer = csv.writer(csv_file)
        tmpList = []
        for key in dict_values.keys():
            tmpList.append(dict_values[key])
        writer.writerow(tmpList)
        print(tmpList)


