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

def downloadExcelFile(filename):
    file_url = "https://www.arcgis.com/sharing/rest/content/items/b5e7488e117749c19881cce45db13f7e/data"
    r = requests.get(file_url, stream = True)
    with open(filename, "wb") as excelFile:
        for chunk in r.iter_content(chunk_size=1024):
            if chunk:
                excelFile.write(chunk)

def createGraphs(name, values, dates, plottype = "Bar"):
    generateTrendprint = False
    calculateStatistics = False

    width = np.diff(dates).min()

    fig, ax = plt.subplots()
    if plottype == "Bar":
        ax.bar(dates, values, align='edge', width=width)
    elif plottype == "Scatter":
        ax.scatter(dates, values)
    elif plottype == "Stack":
        ax.stackplot(dates, values)
    else:
        ax.plot(dates, values)

    # Statistics
    if calculateStatistics:
        spector_data = sm.datasets.spector.load(as_pandas=False)
        spector_data.exog = sm.add_constant(spector_data.exog, prepend=False)
        model = sm.OLS(spector_data.endog, spector_data.exog)
        result = model.fit()
        print(result.summary())

    # Normal print
    ax.bar(dates, values, align='edge', width=width)
    ax.xaxis_date()
    ax.set_title("Covid 19 Fall i {}".format(name))
    ax.set_xlabel("Datum")
    ax.set_ylabel("Antal fall")
    fig.autofmt_xdate()
    plt.savefig("plots/{}.png".format(name))
    plt.close()

    # Trend print
    if generateTrendprint:
        decomposition = sm.tsa.seasonal_decompose(values, model='additive', period=1)
        decomposition.plot()
        plt.savefig('test2.png')

if __name__ == "__main__":
    filename = "Folkhalsomyndigheten_Covid19.xlsx"
    downloadExcelFile(filename)

    selectedSheet = "Antal per dag region"
    df = pd.ExcelFile(filename).parse(selectedSheet)

    dates = df.Statistikdatum

    [createGraphs(list(df)[i], df.iloc[:,i], dates) for i in range(1, len(list(df)))]
