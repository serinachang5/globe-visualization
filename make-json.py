from openpyxl import load_workbook
from random import randrange
import json

def load_data():
    wb = load_workbook("chapter6_a.xlsx")
    table61 = wb["Table 6.1"]
    firstline = 7
    lastline = 227
    data = {"E":[], "F":[], "G":[], "H":[], "I":[], "J":[], "K":[], "L":[]}
    for row in range(firstline, lastline):
        countrycode = int(table61["C"+str(row)].value)
        for col in data:
            try:
                stat = float(table61[col+str(row)].value)
            except ValueError: #no info for this country on this statistic
                stat = None
            data[col].append([countrycode, stat])
    return data

def eval_data(data):
    min_missing = len(data["E"])
    best_col = None
    for col in data:
        missing = 0
        for tuple in data[col]:
            if tuple[1] is None:
                missing += 1
        if missing < min_missing:
            min_missing = missing
            best_col = col
    return best_col

def load_conversion():
    f = open("countries_codes_and_coordinates.csv", 'rb')
    conversions = {}
    f.readline() #skip first line
    line = f.readline()
    while len(line) > 0:
        line = line.rstrip("\"\n")
        cells = line.split("\", \"")
        code = int(cells[3])
        lat = float(cells[4])
        long = float(cells[5])
        conversions[int(code)] = [lat, long]
        line = f.readline()
    return conversions

def check(jsondata):
    print "Length:",len(jsondata)
    datacount = len(jsondata)/3
    randidx = randrange(0,datacount)
    latidx = randidx*3
    longidx = latidx+1
    magidx = latidx+2
    print jsondata[latidx], jsondata[longidx], jsondata[magidx]

def main():
    data = load_data()
    fullname = {"E":"Physical Violence - All Perpetrators - Lifetime", "F":"Physical Violence - All Perpetrators - Last 12 months",
                "G":"Physical Violence - Intimate Partner - Lifetime", "H":"Physical Violence - Intimate Partner - Last 12 months",
                "I":"Sexual Violence - All Perpetrators - Lifetime", "J":"Sexual Violence - All Perpetrators - Last 12 months",
                "K":"Sexual Violence - Intimate Partner - Lifetime", "L":"Sexual Violence - Intimate Partner - Last 12 months"}
    best_col = eval_data(data)
    print "Most data available: column {0} ({1})".format(best_col, fullname[best_col])

    conversions = load_conversion()
    jsondata = []
    for country in data[best_col]:
        code = country[0]
        mag = country[1]
        if mag: #not None aka we have the stat
            mag /= 100.0 #normalize
            loc = conversions[code]
            jsondata.extend([loc[0],loc[1],mag])
    check(jsondata)

    with open("mysearch.json","w") as outfile:
        json.dump([best_col, jsondata], outfile)
        outfile.close()

main()