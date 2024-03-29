#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime as dt
import pandas as pd
import numpy as np
import xlrd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import statsmodels.api as sm
import requests
import os
import time
import sys
import tqdm

CONST_MAX_AGE_OF_DATA_FILE_IN_MINUTES = 60
CONST_DATA_FILE_URI = "https://www.arcgis.com/sharing/rest/content/items/b5e7488e117749c19881cce45db13f7e/data"

def downloadExcelFile(filename):
    if os.path.exists(filename):
        fileLastUpdatedTime = os.stat(filename).st_mtime
        ageOfFileInMinutes = (time.time() - fileLastUpdatedTime) / 60
        if ageOfFileInMinutes < CONST_MAX_AGE_OF_DATA_FILE_IN_MINUTES:
            return
    r = requests.get(CONST_DATA_FILE_URI, stream = True)
    with open(filename, "wb") as excelFile:
        for chunk in r.iter_content(chunk_size=1024):
            if chunk:
                excelFile.write(chunk)

def createGraphs(name, values, dates, plottype = "Plot"):
    generateTrendprint = False
    calculateStatistics = False
    width = np.diff(dates).min()
    title = name

    if name == "Totalt_antal_fall":
        title = "Sverige"

    fig, ax = plt.subplots()
    if plottype == "Bar":
        ax.bar(dates, values, align='edge', width=width)
    elif plottype == "Scatter":
        ax.scatter(dates, values)
    elif plottype == "Stack":
        ax.stackplot(dates, values)
    elif plottype == "Plot":
        ax.plot(dates, values)

    # Rolling average
    df = pd.DataFrame(values)
    df['cum_sum'] = df[name].cumsum()
    df['count'] = range(1,len(df[name])+1)
    df['mov_avg'] = df['cum_sum'] / df['count']
    df['rolling_mean2'] = df[name].rolling(window=7).mean()
    ax.plot(dates, df['rolling_mean2'], "r--")

    # Normal print
    ax.xaxis_date()
    ax.set_title("Nya Covid 19 Fall i {}".format(title))
    ax.set_xlabel("Datum")
    ax.set_ylabel("Antal nya fall")
    ax.legend(["Antal nya fall per dag", "Rullande medeltal (Vecka)"])
    fig.autofmt_xdate()
    plt.savefig("plots/{}.png".format(name))
    plt.close()

    # Trend print
    if generateTrendprint:
        decomposition = sm.tsa.seasonal_decompose(values, model='additive', period=1)
        decomposition.plot()
        plt.savefig("plots/trend-{}.png".format(name))

    # Statistics
    if calculateStatistics:
        spector_data = sm.datasets.spector.load(as_pandas=False)
        spector_data.exog = sm.add_constant(spector_data.exog, prepend=False)
        model = sm.OLS(spector_data.endog, spector_data.exog)
        result = model.fit()
        print(result.summary())

    # Total print
    fig, ax = plt.subplots()
    ax.plot(dates, df['cum_sum'])
    ax.xaxis_date()
    ax.set_title("Totala antalet Covid 19 Fall i {}".format(title))
    ax.set_xlabel("Datum")
    ax.set_ylabel("Totalt antal fall")
    ax.legend(["Antal fall"])
    fig.autofmt_xdate()
    plt.savefig("plots/{}-totalt.png".format(name))
    plt.close()


def numberOfCasesPerDayAndRegion(start_date, end_date):
    selectedSheet = "Antal per dag region"
    df = pd.ExcelFile(filename).parse(selectedSheet)
    dates = pd.to_datetime(df.Statistikdatum)

    mask = (dates > start_date) & (dates <= end_date)
    df = df.where(mask)

    [createGraphs(list(df)[i], df.iloc[:,i], dates) for i in range(1, len(list(df)))]

def numberOfDeaths(start_date, end_date):
    selectedSheet = "Antal avlidna per dag"
    df = pd.ExcelFile(filename).parse(selectedSheet)

    dates = pd.to_datetime(df.Datum_avliden, errors='coerce')

    mask = (dates > start_date) & (dates <= end_date)
    df = df.where(mask)

    values = df.Antal_avlidna
    cumsum = values.cumsum()

    fig, ax = plt.subplots()
    ax.plot(dates, cumsum)
    ax.xaxis_date()
    ax.set_title("Totalt antal avlidna Covid 19")
    ax.set_xlabel("Datum")
    ax.set_ylabel("Antal avlidna")
    ax.legend(["Antal avlidna"])
    fig.autofmt_xdate()
    plt.savefig("plots/Antal_totalt_avlidna.png")

    fig, ax = plt.subplots()
    ax.semilogy(dates, cumsum)
    ax.grid()
    ax.xaxis_date()
    ax.set_title("Totalt antal avlidna Covid 19")
    ax.set_xlabel("Datum")
    ax.set_ylabel("Antal avlidna")
    ax.legend(["Antal avlidna"])
    fig.autofmt_xdate()
    plt.savefig("plots/Antal_totalt_avlidna_log.png")

def numberOfDeathsPerDay(start_date, end_date):
    selectedSheet = "Antal avlidna per dag"
    df = pd.ExcelFile(filename).parse(selectedSheet)
    df.Datum_avliden = pd.to_datetime(df.Datum_avliden, errors='coerce')
    df = df.dropna(subset=['Datum_avliden'])

    dates = pd.to_datetime(df.Datum_avliden, errors='coerce')

    mask = (dates > start_date) & (dates <= end_date)
    df = df.where(mask)

    dates = df.Datum_avliden
    values = df.Antal_avlidna

    # Rolling average
    df = pd.DataFrame(values)
    df['cum_sum'] = df.Antal_avlidna.cumsum()
    df['count'] = range(1,len(df.Antal_avlidna)+1)
    df['mov_avg'] = df['cum_sum'] / df['count']
    df['rolling_mean2'] = df.Antal_avlidna.rolling(window=7).mean()

    fig, ax = plt.subplots()
    ax.plot(dates, values)
    ax.plot(dates, df['rolling_mean2'], "r--")
    ax.xaxis_date()
    ax.set_title("Antal avlidna per dag i Covid 19")
    ax.set_xlabel("Datum")
    ax.set_ylabel("Antal avlidna per dag")
    ax.legend(["Antal avlidna", "Rullande medeltal (Vecka)"])
    fig.autofmt_xdate()
    plt.savefig("plots/Antal_avlidna.png")

def numberOfICECasesPerDay(start_date, end_date):
    selectedSheet = "Antal intensivvårdade per dag"
    df = pd.ExcelFile(filename).parse(selectedSheet)
    df.Datum_vårdstart = pd.to_datetime(df.Datum_vårdstart, errors='coerce')
    df = df.dropna(subset=['Datum_vårdstart'])

    dates = pd.to_datetime(df.Datum_vårdstart, errors='coerce')

    mask = (dates > start_date) & (dates <= end_date)
    df = df.where(mask)

    #dates = df.Datum_vårdstart
    values = df.Antal_intensivvårdade

    # Rolling average
    df = pd.DataFrame(values)
    df['cum_sum'] = df.Antal_intensivvårdade.cumsum()
    df['count'] = range(1,len(df.Antal_intensivvårdade)+1)
    df['mov_avg'] = df['cum_sum'] / df['count']
    df['rolling_mean2'] = df.Antal_intensivvårdade.rolling(window=7).mean()

    # Per day
    fig, ax = plt.subplots()
    ax.plot(dates, values)
    ax.plot(dates, df['rolling_mean2'], "r--")
    ax.xaxis_date()
    ax.set_title("Antal nya intensivvårdade i Covid 19")
    ax.set_xlabel("Datum")
    ax.set_ylabel("Antal nya vårdade")
    ax.legend(["Antal nya intensivvårdade", "Rullande medeltal (Vecka)"])
    fig.autofmt_xdate()
    plt.savefig("plots/Antal_intensivvårdade.png")

    # Totals
    fig, ax = plt.subplots()
    ax.plot(dates, df['cum_sum'])
    ax.xaxis_date()
    ax.set_title("Total antal intensivvårdade i Covid 19")
    ax.set_xlabel("Datum")
    ax.set_ylabel("Totalt antal vårdade")
    ax.legend(["Totalt antal intensivvårdade"])
    fig.autofmt_xdate()
    plt.savefig("plots/Antal_intensivvårdade-totalt.png")

def doit():
    start_date = dt.datetime.now() - dt.timedelta(days=10) - dt.timedelta(days=365) #'2021-01-01'
    end_date =  dt.datetime.now() - dt.timedelta(days=10)  #'2021-03-30'
    downloadExcelFile(filename)
    numberOfCasesPerDayAndRegion(start_date, end_date)
    numberOfDeathsPerDay(start_date, end_date)
    numberOfICECasesPerDay(start_date, end_date)
    numberOfDeaths(start_date, end_date)
    yield

if __name__ == "__main__":
    filename = "Folkhalsomyndigheten_Covid19.xlsx"

    animation = "|/-\\"
    idx = 0
    while doit():
        print(animation[idx % len(animation)], end="\r")
        idx += 1
        time.sleep(0.1)

    #text = 'Working...'
    #print(text)

    #sys.stdout.write("\033[F")
    #for c in text:
    #    sys.stdout.write('\b')
    #sys.stdout.flush()
    #print('Done.')
