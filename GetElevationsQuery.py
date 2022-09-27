
import pandas as pd
import numpy as np
import easygui
import openpyxl


# \Powerbi Tracker.xlsx"
path = easygui.fileopenbox()

df = pd.read_excel(path, engine='openpyxl')

def create_query(dataframe, outfile):
    onboarded = dataframe[dataframe.OnboardingStatus == 'Onboarded']
    filtered = onboarded.dropna(subset=['AzureADId'])
    devices = filtered['AzureADId']
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
    savepath = easygui.filesavebox()
    create_query(df, savepath)