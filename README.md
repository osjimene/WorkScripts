# WorkScripts

Collection of Helpful Work Related Programs

## Introduction

WorkScripts is a collection of Python scripts that can help you with various work-related tasks. These scripts are designed to be easy to use and customizable, so you can adapt them to your specific needs.

## Installation

To use WorkScripts, you'll need to have Python 3 installed on your computer. You can download Python 3 from the official website: https://www.python.org/downloads/

Once you have Python 3 installed, you can download the WorkScripts repository from GitHub:

`git clone https://github.com/yourusername/WorkScripts.git`

After downloading the repository, you can install any dependencies by running:

`pip install -r requirements.txt`

For Authentication, Make sure you have installed the Az.Accounts Powershell module to get your Access Token when making API calls. 

Installation Command:
`Install-Module -Name Az.Accounts -Repository PSGallery -Force`

Once the module is installed, you can import it using the following command:

`Import-Module Az.Accounts`

After importing the module, you can use the Connect-AzAccount cmdlet to authenticate with Azure AD and ARM.

## Usage

To use a script, navigate to the directory where the script is located and run:

`python main.py`


Main.py will prompt you to select which script you want to run in the modules folder.

Note: that there is an example file in the "Example Files" folder that can be used with the GetObjectIds.py, GetElevationsQuery.py and NewUser Query.py scripts.