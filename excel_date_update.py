
# coding: utf-8

# In[11]:


# ensure field names in excel file have no extra spaces or special characters
# change everything under excel variables and feature class variables
# log file will be created in same folder as python file is located
# built using python 3.6
# tested using a file geodatabase with non-versioned data

import arcpy
import logging
import os
import pandas as pd
import sys

# excel variables
excel_file = r"C:\Users\jaso9356\Desktop\dev\py\date_update\date_update.xlsx"
sheet_name = "Sheet1"
excel_parcel_field = "Parcel Number"
excel_date_field = "Date"

# feature class variables
feature_class = r"C:\Users\jaso9356\Documents\ArcGIS\Projects\kim\kim.gdb\tracts"
fc_parcel_field = "parcel_num"
fc_date_field = "title_date"

try:
    
    # set up logging
    current_directory = sys.path[0]
    log_file = os.path.join(current_directory,
                            "date_update.log")
    logging.getLogger('').handlers = []
    logging.basicConfig(filename=log_file,
                        level=logging.INFO,
                        format='%(asctime)s | %(levelname)s | %(message)s',
                        datefmt='%m/%d/%Y %I:%M:%S %p')
    
    # start message
    logging.info("starting update process...")
    
    # read excel file
    df = pd.read_excel(excel_file,
                       sheetname=sheet_name,
                       converters={excel_date_field: pd.to_datetime})
    
    # check to make sure excel column variables match what's in file
    columns = df.columns.values
    
    if excel_parcel_field not in columns:
        logging.error("{} does not match column names in excel file, exiting...".format(excel_parcel_field))
        sys.exit()
        
    if excel_date_field not in columns:
        logging.error("{} does not match column names in excel file, exiting...".format(excel_date_field))
        sys.exit()
    
    # populate dictionary to hold data
    excel_dictionary = {}
    parcel_numbers_in_excel = []
    for i in df.index:
        excel_dictionary[df[excel_parcel_field][i]] = df[excel_date_field][i].date()
        parcel_numbers_in_excel.append(df[excel_parcel_field][i])
        
    # variables to hold process results
    updated_parcel_numbers = []
    
    # update cursor
    with arcpy.da.UpdateCursor(feature_class, [fc_parcel_field, fc_date_field]) as cursor:
        for row in cursor:
            if row[0] in excel_dictionary:
                row[1] = excel_dictionary[row[0]]
                logging.info("{} updated from {} to {}".format(row[0],
                                                                  row[1],
                                                                  excel_dictionary[row[0]]))
                updated_parcel_numbers.append(row[0])
                cursor.updateRow(row)
                
            else:
                logging.info("{} not found in excel spreadsheet".format(row[0]))
    
    # check to see if any parcels in excel did not get processed
    unprocessed_excel_parcels = list(set(parcel_numbers_in_excel) - set(updated_parcel_numbers))
    if len(unprocessed_excel_parcels) > 0:
        parcels_string =  ",".join(unprocessed_excel_parcels)
        logging.info("parcels that did not get updated: {}".format(parcels_string))
    
except Exception as e:
    logging.error(e)

