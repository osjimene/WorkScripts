
import pandas as pd
import numpy as np
import easygui
import openpyxl


##########################################################################################
# This script will take a list of AzureADIds and return a prepopulated Copy and paste query for Kusto.
# This cuts down time having to modify and create a query and send to SAW for Advanced Hunting queries.
# ESPECIALLY useful for when you have to use thousands of devices and include them in a KQL query. 

##########################################################################################

# \Powerbi Tracker.xlsx"
path = easygui.fileopenbox()

df = pd.read_excel(path, engine='openpyxl')

def create_query(dataframe, outfile):
    devices = dataframe['AzureADId']
    list_query = devices.tolist()
    with open(outfile, 'w') as output:
        output.write(str(f"""let devices =datatable (AadDeviceId : string) {list_query};
let filtered = devices 
| where AadDeviceId in (devices);
let info = materialize(DeviceInfo
| where Timestamp >= ago(30d)
|where AadDeviceId in (filtered)
|summarize arg_max(Timestamp,*) by AadDeviceId
|project Timestamp,DeviceId,AadDeviceId, DeviceName);
let deviceIds = info | project DeviceId;
DeviceProcessEvents
| where Timestamp between (datetime() .. datetime())
| where DeviceId in (deviceIds)
| summarize countif(FileName =='consent.exe') by DeviceId"""))  

if __name__ == "__main__":
    savepath = easygui.filesavebox(filetypes=['*.txt'])
    create_query(df, savepath)