import requests
import pandas as pd
import numpy as np
import easygui
import json
import openpyxl
import subprocess


##########################################################################################
# This script will take a list of AzureADIds and return the corresponding ObjectIds.

# The script will take an excel file with a column titled "AzureADId" and return a new excel file with a column titled "ObjectId"

#Requires Dependency on Connect-AzAccount which is found in the Az.Accounts module.
#You can install the module with Install-Module -Name Az.Accounts -Repository PSGallery -Force via Powershell. 
##########################################################################################



# Get the GraphAPI Key for the calls that you are required to make.
connect = 'powershell.exe -Command "Connect-AzAccount"'
# time.sleep(5) # add a 5 second delay
token = 'powershell.exe cls; (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token'

# Run the command and capture the output
connection = subprocess.check_output(connect, shell=False)
secretkey = subprocess.check_output(token, shell=False).decode('utf-8').strip()

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