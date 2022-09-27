import pandas as pd
import argparse
import easygui


# def parse_args():
#     """ Setup the input and output arguments for the script
#     Return the parsed input and output files
#     """
#     parser = argparse.ArgumentParser(description='Create kusto query')
#     parser.add_argument('infile',
#                         type=argparse.FileType('r'),
#                         help='csv file used as template')
#     parser.add_argument('outfile',
#                         type=argparse.FileType('w'),
#                         help='text file containing the query')
#     return parser.parse_args()

infile = easygui.fileopenbox(msg='Please select a csv file with AadDeviceIds in the first column')

def create_query(infile):
    df = pd.read_csv(infile)
    firstcolumn = df.iloc[:,0]
    list_query = firstcolumn.tolist()
    # with open(outfile, 'w') as output:
        # output.write(str(f"let devices =datatable (AadDeviceId : string) {list_query};\nlet filtered = devices \n| where AadDeviceId in (devices)\nlet info = materialize(DeviceInfo\n| where Timestamp >= ago(30d)\n|where AadDeviceId in (filtered)\n|summarize arg_max(Timestamp,*) by AadDeviceId\n|project Timestamp,DeviceId,AadDeviceId, DeviceName);\nlet deviceIds = info | project DeviceId;\nDeviceProcessEvents\n| where Timestamp >= ago(30d)\n| where DeviceId in (deviceIds)\n| summarize countif(FileName =='consent.exe') by DeviceId"))  
        # output.write(str(f"""let devices = datatable (AadDeviceId: string) {list_query};
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
    # args = parse_args()
    query = create_query(infile)
    easygui.msgbox(msg=query)

