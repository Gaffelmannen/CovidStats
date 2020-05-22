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
    ax.set_title("Covid 19 Fall i {}".format(title))
    ax.set_xlabel("Datum")
    ax.set_ylabel("Antal fall")
    ax.legend(["Antal fall", "Rullande medeltal (Vecka)"])
    fig.autofmt_xdate()
    plt.savefig("plots/{}.png".format(name))
    plt.close()

    # Trend print
    if generateTrendprint:
        decomposition = sm.tsa.seasonal_decompose(values, model='additive', period=1)
        decomposition.plot()
        plt.savefig('test2.png')

    # Statistics
    if calculateStatistics:
        spector_data = sm.datasets.spector.load(as_pandas=False)
        spector_data.exog = sm.add_constant(spector_data.exog, prepend=False)
        model = sm.OLS(spector_data.endog, spector_data.exog)
        result = model.fit()
        print(result.summary())

def numberOfPerDayAndRegion():
    selectedSheet = "Antal per dag region"
    df = pd.ExcelFile(filename).parse(selectedSheet)
    dates = df.Statistikdatum
    [createGraphs(list(df)[i], df.iloc[:,i], dates) for i in range(1, len(list(df)))]

def numberOfDeathsPerDay():
    selectedSheet = "Antal avlidna per dag"
    df = pd.ExcelFile(filename).parse(selectedSheet)
    df.Datum_avliden = pd.to_datetime(df.Datum_avliden, errors='coerce')
    df = df.dropna(subset=['Datum_avliden'])

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
    ax.set_title("Antal avlidna i Covid 19")
    ax.set_xlabel("Datum")
    ax.set_ylabel("Antal fall")
    ax.legend(["Antal avlidna", "Rullande medeltal (Vecka)"])
    fig.autofmt_xdate()
    plt.savefig("plots/Antal_avlidna.png")

if __name__ == "__main__":
    filename = "Folkhalsomyndigheten_Covid19.xlsx"
    downloadExcelFile(filename)
    numberOfPerDayAndRegion()
    numberOfDeathsPerDay()
