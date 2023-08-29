import easygui as eg
import subprocess
import os

# Get the path of the Modules folder
modules_folder = os.path.join(os.path.dirname(__file__), "Modules")

# Dictionary that maps the choices in eg. choicebox to the python scripts in this program.
scripts = {
    "Get ObjectIds from GraphAPI": "GetObjectIds.py",
    "Get Elevations Query for Kusto": "GetElevationsQuery.py",
    "Comparison Script": "Comparison Script.py",
    "Query for New Standard Users": "NewUser Query.py",
    "Query to add Azure Subscriptions to ServiceTree": "ServiceTreeAPI.py"
}

# Show the choicebox and get the user's selection
selection = eg.choicebox(msg="What Module/Script would you like to run?", title="Helpful WorkScripts", choices=list(scripts.keys()))

# Get the filename corresponding to the user's selection from the scripts dictionary
filename = os.path.join(modules_folder, scripts[selection])

# Run the selected script using subprocess
subprocess.run(["python", filename])


