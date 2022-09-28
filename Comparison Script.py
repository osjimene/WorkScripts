import pandas as pd
from openpyxl import load_workbook
import hashlib
import ipaddress
from netaddr import IPNetwork, IPAddress
import easygui


#set path to the document that you want to make the comparison with.
# path = r"C:\Users\osjimene\OneDrive - Microsoft\NDR Magellan\NDR PowerBI\Alan Turing Meeting Rooms 27July2020.xlsx"
path = easygui.fileopenbox(msg="Please select a File" ,title="Open Excel File")


#set up the worksheets as dataframes ie. df(Collector) df2(Baseline) df3(Subnets)
df = pd.read_excel(path, engine="openpyxl", sheet_name = 0)
df2 = pd.read_excel(path, engine="openpyxl", sheet_name =1)
df3 = pd.read_excel(path, engine="openpyxl", sheet_name = 2)


#this takes the thrid tab with subnet information and returns the matching VLAN segment for
#the devices. 
#(Based on Men and Mice tool data pulled from OneAsset)
subnets = pd.Series(df3.Title.values, index = df3.RangeTitle).to_dict()

IPAddress = df.IPAddress.tolist()

ipVlan = {}

for i in subnets:
    for j in IPAddress:
        if ipaddress.ip_address(j) in ipaddress.ip_network(i):
            ipVlan[j] = subnets[i] 


ipVlan

subnetList = pd.DataFrame(list(ipVlan.items()),columns = ['IPAddress','VLAN'])

df = df.merge(subnetList, on='IPAddress')


#////////////////////////////////Defined Functions///////////////////////////////////////

# set function that will take away all the double occurences of MAC addresses. 


def removeDoubleMac(DataFrame):
    DataFrame = DataFrame.drop_duplicates(subset=['MacAddress'],keep='first')
    
    return DataFrame



#return MAC address values to Hashed values in case NDR displays them this way. 

def macHash(x):
    dfmac = x['Mac'].astype(str)
    scrubbedMac = []
    
    for mac in dfmac:
        oui = mac[0:8].upper()
        
        hashObject = hashlib.sha1(mac[8:].encode().lower())
        hexDig = hashObject.hexdigest()
        
        if len(mac) == 17:
            scrubbed = "MacPII_"+oui+"_"+hexDig
            scrubbedMac.append(scrubbed)
            
        else:
            scrubbed = "Null"
            scrubbedMac.append(scrubbed)
        
    x['MacAddress']= scrubbedMac
    return x

#Retun the MAC address in the correct : format to compare the 2 values.

def macFormat(x):
    dfformat = x['MacAddress'].astype(str)
    formattedMac = []
    
    for mac in dfformat:
        if len(mac) == 12:
            scrubbed = mac[0:2]+":"+mac[2:4]+":"+mac[4:6]+":"+mac[6:8]+":"+mac[8:10]+":"+mac[10:12]
            formattedMac.append(scrubbed.upper())

        elif len(mac) == 17:
            scrubbed_2 = mac[0:2]+":"+mac[3:5]+":"+mac[6:8]+":"+mac[9:11]+":"+mac[12:14]+":"+mac[15:17]
            formattedMac.append(scrubbed_2.upper())

        else:
            formattedMac.append(mac)
    x['MacAddress'] = formattedMac
    return x


            


#Get the total number of devices that were discovered by each device (Based off MAC address)

def totalCount(DataFrame):
    print(DataFrame['MacAddress'].count())
    
    
#Get comparison numbers. How many matched on collector/how many matched on Universe, etc.
#Had some issues with combining different named columns, make sure that the columns that you are merging are named the same. 
match = []
collector_match = []
universe_match = []

def collectorMatch(CollectorDF, RepoDF):
    
    match = CollectorDF.merge(RepoDF, on='MacAddress')
                
    return match

def repoMatch(CollectorDF,RepoDF):
    
    match = RepoDF.merge(CollectorDF, on ='MacAddress')
    
    return match
    
def comparison(CollectorDF, RepoDF):
    
    CollectorDF=CollectorDF.add_prefix('Collector_')
    RepoDF=RepoDF.add_prefix('Repo_')
    
    match = CollectorDF.merge(RepoDF, left_on='Collector_MacAddress',right_on='Repo_MacAddress', how='inner')
    
    return match

def collectorNoMatch(CollectorDF, RepoDF):
    match = CollectorDF.merge(RepoDF, on='MacAddress', how='left', indicator=True).query('_merge == "left_only"').drop('_merge', 1)
    
    return match

def repoNoMatch(CollectorDF, RepoDF):
     match = CollectorDF.merge(RepoDF, on='MacAddress', how='right', indicator=True).query('_merge == "right_only"').drop('_merge', 1)
     
     return match


def collectorTotal(CollectorDF):
    collectorTotal = CollectorDF['MacAddress'].count()
    return collectorTotal

def repoTotal(RepoDF):
    repoTotal = RepoDF['MacAddress'].count()
    return repoTotal



#////////////////////////////////////Working Script////////////////////////////////////////

# make sure that the MAC column reads MacAddress (df.rename(columns = {'MAC':'MacAddress'},inplace = True))
df2.rename(columns = {'MAC':'MacAddress'},inplace = True)
#Make sure that the Mac Addresses are in the same format (macFormat(x))
macFormat(df)
macFormat(df2)

print('Data Formatted')


#Make sure that you remove any double occurences on the MAC address (removeDoubleMac(DF))
df= removeDoubleMac(df)
df2= removeDoubleMac(df2)

print("Doubles Removed")

#Now you can run the data and pull comparisons.
#Compare the two datasets side by side on matching MAC values (comparison(CollectorDF, RepoDF))

comparison =comparison(df,df2)

print("Comparison Complete")


#Compare the two datasets and get the values that didnt match on the Collector(collectorNoMatch(CollectorDF,RepoDF)
collectorNoMatch = collectorNoMatch(df,df2)
collectorNoMatchTotal = collectorNoMatch.count()

collectorNoMatch = collectorNoMatch.reset_index(drop = True)


print('No-Match Collector Comparison Complete')

#Compare the two datasets and get the values that didnt ... (repoNoMatch(CollectorDF, RepoDF)
repoNoMatch = repoNoMatch(df,df2)
repoNoMatchTotal = repoNoMatch.count()

repoNoMatch = repoNoMatch.reset_index(drop = True)

repoNoMatch = repoNoMatch.drop(repoNoMatch.columns[[0,2,3,4,5,6,7,8,9]],axis = 1)

#repoNoMatch = repoNoMatch[['MacAddress','IPAddress']]

repoNoMatch

print('No-Match Repo Comparison Complete')

#Compare the two datasets and get the values that matched in that dataset(collectorMatch(CollectorDF, RepoDF))
#Compare the two ...(repoMatch(CollectorDF, RepoDF))
collectorMatch = collectorMatch(df,df2)
collectorMatchTotal = collectorMatch.count()
repoMatch =repoMatch(df,df2)
repoMatchTotal = repoMatch.count()

collectorMatch = collectorMatch.reset_index()
collectorMatch = collectorMatch.drop(['index'],axis = 1)
repoMatch = repoMatch.reset_index()

print('Total Number Calculations Complete')



#Get total number of Devices on each list by MAC Address collectorTotal(CollectorDF) or repoTotal(RepoDF):
collectorTotal = collectorTotal(df)

repoTotal = repoTotal(df2)

print('Total Numbers Counts Complete')



#This final section will save and complete it to a excel file, not needed for PowerBi

#Set conditions to use Openpyxl module to write a new sheet
#into the excel when finished. 
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine = "openpyxl")
writer.book = book



#These are here for labeling purposes only. 
collector = pd.DataFrame({"Collector"})

repo = pd.DataFrame({'Repo'})

#This will be at the end to compile all totals and append on last 
# worksheet

df4 = pd.DataFrame(
{"Collector Totals" : [collectorTotal],
"Repo Totals" : [repoTotal],
 "Collector Match" : [collectorMatch['MacAddress'].count()],
 "Repo Match" : [repoMatch['MacAddress'].count()],
 "Collector No-Match": [collectorNoMatch['MacAddress'].count()],
 "Repo No-Match" : [repoNoMatch['MacAddress'].count()]})

df4.to_excel(writer, sheet_name = 'Match Totals')

collectorMatch.to_excel(writer, sheet_name='Match Totals', startrow =5, startcol=0)

collectorNoMatch.to_excel(writer, sheet_name = 'CollectorNoMatch',startrow = 0, startcol = 0)

repoNoMatch.to_excel(writer, sheet_name = 'RepoNoMatch',startrow = 0, startcol = 0)




writer.save()
writer.close()







    
        
