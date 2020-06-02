# SharePoint List Data Import Tool
SharePoint List Data Import Tool is an easy-to-use .NET application to retrieve the SharePoint list data into database or excel format. It helps user to export Share Point data as well as sync data upon the click of a button. This tool was born out of business needs:

1) Client maintains Dimension/Master data into SharePoint list which is easy to maintain but to use in Cube it must be loaded into database. It require good amout of effort to load data. This tool is reducing the effort in a single click. 

2) Developer can learn from ready solution and can enhance the application for thier requirement. It helps to reduce the overall effort.

3) To sync Share point data into SQL database can achive by one click.

# How to Install:

Performt the following steps tp deploy the application in IIS 

STEP 1: Right click on the sites and click on Add web sites button.
STEP 2: Give the site name and provide the physical path of the application as store in the system.
STEP 3: Click on the OK button to create the site.
STEP 4: Right click on the deployed application and browse the application.

# How to use it:

1) Open the application in a browser. Enter the SharePoint link in the text box and click on the Display list items button. The button click event will fetch the list item names and bind to the drop down list.

2) The web application will display all the SharePoint lists in the drop down list based on the SharePoint link given in the text box

3) Select the SharePoint ListItem name in the drop down list which we want to store the  require the data to  our local system .

4) Select the Format which is required in the radio button list and click on the Export button.The button event will store the data to local system based on the selection of the radio button.

5) If Excel radio button is selected then List item data will be stored to Excel with list item name and if we select the sql database then we have to provide the sql database server details based on the way we are connecting to the data base.

The Authentication is windows then we have to provide the below two details
a. Datasource 
b. Database

If the Authentication is sql server then we have to provide all the below details
a.Data source
b.DataBase Name
c.UserId
d.Password

6) Click on export button, it will create the table in database with list name and save the data to that table.

# Solution :
To fetch the SharePoint list with clientcontext object. The following assembles are required and already available within the source code :
a) Microsoft.SharePoint.Client.dll and 
b) Microsoft.SharePoint.Client.Runtime.dll

This object is responsible for connecting to the server and sending all queued commands for processing.

ClientContext.ExecuteQuery method triggers the communication between the SharePoint Site and web application. Once the action performed store the data into ListItem collection and perform the required operations (Storing the data to database, Excel) on list  item.

# Prerequisites

1. The user must have the admin access on the SharePoint url to fetch the lists and list details
2. The user must have access on the database to execute DDL and DML commands
3. .NET Framework 4.5 must be installed in the system
4. IIS 6.0 or above





