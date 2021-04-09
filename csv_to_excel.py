# -*- coding: utf-8 -*-
"""
Created on Tue Mar  2 11:26:32 2021

@author: karlo
"""


def install(package):
    pip.main(['install', package])

try:
    import pandas as pd
    import numpy as np
    import os
except ModuleNotFoundError:
    import pip
    print("Need to install packages.. Please wait..")
    # or
    install("pandas") # the install function from the question
    install ("numpy")
    install("os")
    print ("Required packages installed.")

def csv_to_excel(file_name):
    df = pd.read_csv(file_name, sep=";", skiprows=2)
    emv = df[" mV"]
    return emv
def get_vol(file_name):
    df = pd.read_csv(file_name, sep=";", skiprows=2)
    vol = df["Total Volume"]    
    return vol

def calculate_(name, dataframe):
    conc = float(name[-6])*10**(-1*float(name[-4]))

    v0 = 25.0
    z_cat = 1
    z_an = -1
    magzb = 1
    # calculating concentrations etc.
    dataframe["c(mol/L)"] = conc * (dataframe["Total Volume"] / (v0+dataframe["Total Volume"]))
    dataframe["log c"] = np.log10(dataframe["c(mol/L)"])
    dataframe["I"] = 0.5 * ((dataframe["c(mol/L)"] * magzb * z_cat**2)+ (dataframe["c(mol/L)"] * z_cat * z_an**2))
    dataframe["log I"] = np.log10(dataframe["I"])
    dataframe["log fI"] = -0.512 * z_an**2 * ((np.sqrt(dataframe["I"])) /  (1+(np.sqrt(dataframe["I"]))-0.2*dataframe["I"]  ))
    dataframe["fI"] = 10**dataframe["log fI"]
    dataframe["aI"] = dataframe["c(mol/L)"] * dataframe["fI"]
    dataframe["log (aI)"] = np.log10(dataframe["aI"])
    return dataframe


def chart_conf_dict(sheet): # defining charts
    chart_conf_dict = {"name":       "mV",
                       "line": {"none": True},
                       "marker": {"type": "square",
                                  "size": 6,
                                  "border": {'color': 'blue'},
                                  "fill":   {'none': True}},
                       'categories': '=' + sheet +'!$k$3:$k$27', 
                       'values':     '=' + sheet + '!$c$3:$c$27'}
    return chart_conf_dict

def chart_conf_dict_summary(sheet): # definign summary chart
    chart_conf_dict = {"name":      sheet[-6:], 
                       'categories': '=' + sheet +'!$k$3:$k$27', 
                       'values':     '=' + sheet + '!$c$3:$c$27'}
    return chart_conf_dict


if __name__ == "__main__":
    curr_dir = os.getcwd()
    os.chdir(curr_dir)
    list_files = list()
    writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter') 
    workbook  = writer.book
    for file in os.listdir(curr_dir):
        if file.endswith(".csv"):
            os.rename(file, file.replace(' ', '_')) # replaces spaces in filenames with underscores
            list_files.append(file.upper()) # adds filenames to the list, capitalizing
    dict_files = {file[:-4]:csv_to_excel(file) for file in list_files} 
    # creating a dictionary, key = filename, value = dataframe
        
    for file in list_files: # getting the volume
        vol = get_vol(file)
        break
    
    for k, v in dict_files.items():
        k_fp = k.split('_RES', 1)[0] # 1st part of the filename, before response
        k_sp = k.split('_RES', 1)[1] # 2nd part of the filename, after response
        k_sheet = k_fp + "_" + k_sp[-6:] # sheet names for the excel file
        
        dict_files[k] = pd.concat([vol, v], axis = 1)
        df = calculate_(k, dict_files[k])
        df.to_excel(writer, sheet_name=k_sheet)
        worksheet = writer.sheets[k_sheet] # writing the sheets in the new excel file
        
        # defining the charts in each sheet
        chart1 = workbook.add_chart({"type": "scatter", "subtype": "smooth"}) 
        chart1.add_series(chart_conf_dict(k_sheet))
        chart1.set_title({ 'name': k,
                           'name_font': {'size': 14, 'bold': True},})
        chart1.set_x_axis({"name": "log a(DDS)",
                           'min': -8, 'max': -2,
                           'name_font': {'size': 9, 'bold': False},})
        chart1.set_y_axis({"name": "E/mV",
                            'major_gridlines': {
                            'visible': False,},
                            'name_font': {'size': 9, 'bold': False},})

        chart1.set_legend({'none': True})

        worksheet.insert_chart('M4', chart1)
    
    worksheet2 = workbook.add_worksheet("Summary") # creating the summary sheet
    chart2 = workbook.add_chart({"type": "scatter", "subtype": "smooth_with_markers"}) 
    
    for sheet in dict_files.keys(): # adding lines to summary chart
        sheet_fp = sheet.split('_RES', 1)[0] # 1st part of the filename, before response
        sheet_sp = sheet.split('_RES', 1)[1] # 2nd part of the filename, after response
        k_sheet = sheet_fp + "_" + sheet_sp[-6:] # sheet names for the excel file
        chart2.add_series(chart_conf_dict_summary(k_sheet)) # adding the chart
        
    # definign the summary chart details   
    chart2.set_title({'name': next(iter(dict_files))[:-7]}) # the name is given accorinf to first filename
    chart2.set_size({'width': 1400, 'height': 720})
    chart2.set_y_axis({'major_gridlines': {
                       'visible': False,},})
    chart2.set_x_axis({'min': -8, 'max': -3})
    worksheet2.insert_chart('B2', chart2)

    writer.save()
    writer.close()
    workbook.close()


    
    
    
    
    
    
    
    
    