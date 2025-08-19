# Apps-Script-Mail-and-Certificate-generator

Certificate & Email Generator

This is a comprehensive Google Apps Script solution designed to automate the generation and distribution of personalized certificates and emails. It streamlines a manual, repetitive process into a seamless, user-friendly workflow directly within a Google Sheet.

Features

•	Automated Certificate Generation: Collects participant data (Name, Email, Course) from an Excel sheet and automatically generates personalized certificates from a Google Docs template.

•	PDF Conversion & Archiving: Converts each generated certificate into a PDF file and automatically saves it to Google Drive, with the link populated back to the Google Sheet.

•	Customizable Email Templates: Features a user-friendly sidebar UI that allows for the creation and management of custom email templates to be used for mass communication.

•	Targeted Email Distribution: Sends personalized emails to recipients with the option to attach their generated certificate directly from the sidebar.

•	Interactive Sidebar UI: A custom HTML sidebar provides a clear interface for users to select actions, manage settings, and track the status of certificate generation and email delivery.

•	Status Tracking: Updates a dedicated "Status" column in the spreadsheet with "Done" and includes a direct link to the generated certificate.

Technology Stack

•	Google Apps Script (GAS): The core backend scripting language for all automation and data manipulation.

•	Google Sheets API: Used to read data from the main sheet and update the status column.

•	Google Docs API: Facilitates the creation of new documents from a template and the population of data fields.

•	Google Drive API: Used to save the generated PDF certificates and manage file links.

•	Gmail API: For sending personalized emails to recipients with attachments.

•	HTML, CSS, & JavaScript: Used to build the custom sidebar user interface, providing a front-end experience within the Google Sheet environment.

Visual Walkthrough

Since this project runs within a private Google environment, it cannot be hosted on a public website. Here is a link to the case study showing a visual walkthrough demonstrating the system's key functionality.

How to Use

This project is intended to be used by copying the script into a Google Sheet.

1.	Create a New Google Sheet: Go to to create a new spreadsheet for your data.
2.	Open Apps Script: In the Google Sheet, navigate to Extensions > Apps Script to open the script editor.
3.	Copy the Code: Copy the contents of Code.gs and Sidebar.html into your new Apps Script project.
4.	Get Template and Folder IDs:
o	Google Docs Template ID: The Document ID is the long string of characters in the URL of your Google Docs template, found between /d/ and /edit.
o	Google Drive Folder ID: The Folder ID is the long string of characters at the end of the URL when you open the folder in Google Drive, found after /folders/.
5.	Run the Script: Input the templates ID then ave the project and run the onOpen() function to display the sidebar.
6.	Open Google Sheets: On the previous sheet opened, press the certificate generator tab, then click on open sidebar to the sidebar.

License

This project is licensed under the MIT License - see the LICENSE.md file for details.

