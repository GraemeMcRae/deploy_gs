# How To Create a Google Service Account

If you want to use the gspread Python package to access a Google Sheets spreadsheet, you'll need to create a Google Service Account, and a credentials file called something like google_credentials.json to authenticate using this service account.

## Purpose

This document exists so that:

- You can use deploy_gs.py to deploy Google spreadsheet formulas


## Overview of Steps

- Create a Google Cloud Project
- Enable Google Sheets API and Google Drive API
- Create a Service Account 
- Download credentials
- Share your Google Sheet with the service account


## Create a Google Cloud Project

- Go to Google Cloud Console ( console.cloud.google.com )
- Click on the project dropdown (top left, next to "Google Cloud")  
Note: if you have a previous project, click it, and use the Project Picker.
- Click "New Project"
- Name it: projectname (or whatever you prefer)
- Click "Create"
- Wait for it to be created, then select it


## Enable Google Sheets API

- Navigate the Google Cloud Console starting from the ☰ (Menu) icon (three horizontal bars)
- In your projectname project, go to ☰ > APIs & Services > Library
- Search for: Google Sheets API
- Click on it, then click Enable
- Also search for and enable: Google Drive API (needed for file access)


## Create a Service Account and download credentials

- Go to ☰ > APIs & Services > Credentials
- Click Create Credentials > Service Account
- Fill in: Service account name: service-account-name, Service account ID: (auto-filled), Description: Service account for Project Name to access Google Sheets
- Click Create and Continue
- Click "Continue" to skip "Grant this service account access to project"
- Click "Done" to skip "Grant users access to this service account"

## Download credentials

- On the Credentials page, you should see your service account listed
- Click on the service account email (service-account-name@service-account-id.iam.gserviceaccount.com)
- Go to the Keys tab
- Click Add Key > Create new key
- Choose JSON format
- Click Create
- A JSON file will download - this contains your credentials - rename it to google_credentials.json
- ***CRITICAL:*** Add it to .gitignore so you never commit it to GitHub!


## Share your Google Sheet with the service account

- Go to drive.google.com
- If you are prompted, log in using your Google credentials

   Then, either:

   - Click `+ New` > Google Sheets > Blank spreadsheet

   Or:

   - Open an existing Google spreadsheet

   Then continue with the next step:

- With the spreadsheet open, click the "earth" icon on the Share button at the upper right of the screen
- In the "Add people..." field, enter the full email address of your service account, e.g.  
`service-account-name@service-account-id.iam.gserviceaccount.com`
- To the right of the service account name, select the appropriate level of access, e.g. "Editor"
- Uncheck the "Notify people" box
- Click the blue "Share" button


## Summary

- You have created a Google Cloud Project which you can visit at any time right here:  
`console.cloud.google.com`
- To review these steps, start by selecting the Google Cloud Project you just created, which is identified by a rectangular box just to the right of the words "Google Cloud" at the top of your screen.
- You can navigate within your project by clicking the ☰ (Menu) icon (three horizontal bars near the upper left of the screen).
- Choose ☰ > APIs & Services to see the two APIs you enabled - Google Sheets API and Google Drive API.
- Choose ☰ > APIs & Services > Credentials to see the service account you created
