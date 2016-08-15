# Project CSOM Read Local CustomFields

The github.com/OfficeDev/Project-CSOM-Read-Local-CustomFields sample uses C# and the Project CSOM to demonstrate how to access custom fields that are defined within a project. 

Users typically access local custom fields by opening a project using the Project Professional  or Project Pro for O365, then selecting Custom Fields from the Project tab of the ribbon. 

## Scenario

I want to be able to retrieve project local custom fields so that I can display properties/data unique to my projects.

### Using App

1.	Add the Project CSOM client package [here](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM/)
2.	Update the PWA site
3.	Update the login/password to your PWA site.
4.  Upload the sample Project mpp file to PWA and publish
5.	Run the app

### Prerequisites
To use this code sample, you need the following:
* Project Online tenant.
* Visual Studio 2013 or later 
* Project CSOM client DLL.  It is available as a Nuget Package from [here](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM/)
* The project named "Local Custom Fields" added to the Project PWA site.


## How the sample affects your tenant data
This sample runs CSOM methods that reads the contents of uploaded project "Local Custom Fields" for the specified user. Tenant data will be affected.

## Additional resources
* [Local and Enterprise Custom Fields](https://msdn.microsoft.com/en-us/library/office/ms447495(v=office.14).aspx)

* [ProjectContext class](https://msdn.microsoft.com/en-us/library/office/microsoft.projectserver.client.projectcontext_di_pj14mref.aspx)

* [Client-side object model (CSOM)](https://aka.ms/project-csom-docs)

## Copyright

Copyright (c) 2016 Microsoft. All rights reserved.

