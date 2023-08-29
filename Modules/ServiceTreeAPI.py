import aiohttp
import asyncio
import requests
import pandas as pd
import numpy as np
import easygui
import json
import openpyxl
from urllib.parse import urlencode
import subprocess

##########################################################################################
# This script will take a list of Azure SubscriptionIds and add them to Service tree using the post method. 

# This is useful for when you have to add a large amount of subscriptions to Service Tree.
# The imported file should be a csv with a column titled "Subscriptions" and the subscription ids in the rows below.
##########################################################################################



# Get the GraphAPI Key for the calls that you are required to make.
connect = 'powershell.exe -Command "Connect-AzAccount"'
# time.sleep(5) # add a 5 second delay
token = 'powershell.exe cls; (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token'

# Run the command and capture the output
connection = subprocess.check_output(connect, shell=False)
secretkey = subprocess.check_output(token, shell=False).decode('utf-8').strip()

file = easygui.fileopenbox(msg="Please choose a file to import")
df = pd.read_csv(file)
subs = df['Subscriptions'].tolist()

async def addSubscription(data):
    # Set the endpoint URL
    endpoint = "https://servicetreeprodwest.azurewebsites.net/api/ServiceHierarchy(54bb101d-d647-4089-b71d-99239113cb6c)/ServiceTree.AddAzureSubscription()"
    
    # Set the authorization header
    secret = 'Bearer '+ secretkey
    
    # Set the headers for the HTTP request
    http_headers = {'Authorization' : secret,
                    'ConsistencyLevel': 'eventual',
                    'Content-Type': 'application/json',
                    'Connection':'keep-alive',
                    'Accept-Encoding': 'gzip,deflate,br'}
    
    # Set the JSON body for the HTTP request
    body = {"AzureSubscription":{"AzureCloud":"Public","Environment":"NonProduction","EvictionImpact":"Standard","IsHostedOnBehalfOf":"False","SubscriptionId":f"{data}"}}
    
    # Create an aiohttp ClientSession
    async with aiohttp.ClientSession() as session:
        # Send a POST request to the endpoint with the specified headers and JSON body
        async with session.post(endpoint, headers=http_headers, json=body) as response:
            # Get the JSON response data
            response_data = await response.json()
            # Return the response data
            return response_data

async def main():
    # Set the batch size
    batch_size = 10
    # Calculate the number of batches
    num_batches = len(subs) // batch_size + (len(subs) % batch_size != 0)
    # Process the subscriptions in batches
    for i in range(num_batches):
        # Get the current batch of subscriptions
        batch_subs = subs[i*batch_size:(i+1)*batch_size]
        # Create a list of tasks for the current batch of subscriptions
        tasks = [addSubscription(sub) for sub in batch_subs]
        # Run all tasks concurrently using asyncio.gather
        await asyncio.gather(*tasks)

# Run the main coroutine using asyncio.run
asyncio.run(main())



