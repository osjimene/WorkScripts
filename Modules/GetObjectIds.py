import requests
import pandas as pd
import numpy as np
import easygui
import json
import openpyxl


##########################################################################################
# This script will take a list of AzureADIds and return the corresponding ObjectIds.

# The script will take an excel file with a column titled "AzureADId" and return a new excel file with a column titled "ObjectId"
##########################################################################################



secretkey= easygui.enterbox(msg="Please enter your GraphAPI Key...")
url = "https://graph.microsoft.com/v1.0/devices?"

# path = Powerbi Tracker.xlsx"
path = easygui.fileopenbox(msg="Please select the file you want to import")
df = pd.read_excel(path, engine='openpyxl')

def response(device):
    endpoint = f"https://graph.microsoft.com/v1.0/devices?$select=id,deviceId,displayName&$filter=deviceId%20eq%20%27{device}%27"
    secret = 'Bearer '+ secretkey
    http_headers = {'Authorization' : secret,
                    'ConsistencyLevel': 'eventual'}
    r = requests.get(endpoint, headers = http_headers).json()
    try:
        return r['value'][0]['id']
    except IndexError:
        return "IndexError"
    except KeyError:
        return "KeyError"


df['ObjectId'] = df.apply(lambda x: response(f'{x.AzureADId}'), axis=1)

outfile = easygui.filesavebox(msg="Please choose where you want the outfile saved")

df.to_excel(outfile,engine='openpyxl')