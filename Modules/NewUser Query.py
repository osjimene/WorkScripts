import pandas as pd
import easygui

##########################################################################################
# This script will take a list of AzureADIds and return an Advanced Hunting query for Kusto. 

# The query helps identify users that have not shown many Admin elevation events in the last 30 days.

# This was used to help identify users that could be onboarded to Standard User with low impact to productivity.
##########################################################################################

infile = easygui.fileopenbox(msg='Please select a csv file with AadDeviceIds in the first column')

def create_query(infile):
    df = pd.read_csv(infile)
    firstcolumn = df.iloc[:,0]
    list_query = firstcolumn.tolist()
    query = (f"""let devices = datatable (AadDeviceId: string) {list_query};
    let filtered = devices
    |where AadDeviceId in (devices);
    let info = materialize(DeviceInfo
    |where Timestamp >= ago(30d)
    |where AadDeviceId in (filtered)
    | summarize arg_max(Timestamp,*) by AadDeviceId
    | project Timestamp, DeviceId, AadDeviceId, DeviceName);
    let deviceIds = info
    | project DeviceId;
    DeviceProcessEvents
    | where TimeStamp >= ago(30d)
    | where DeviceId in (deviceIds)
    | summarize countif(FileName == 'consent.exe') by DeviceId""")
    return query


if __name__ == "__main__":
    query = create_query(infile)
    easygui.msgbox(msg=query)

