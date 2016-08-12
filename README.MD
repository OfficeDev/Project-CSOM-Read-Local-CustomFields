# Access CSOM Local Custom Fields

The TaskCustomFields sample uses C# and the Project CSOM to demonstrate how to access custom fields that are defined within a project. 

Users typically access local custom fields by opening a project using the Project desktop application or Project 365, then selecting Custom Fields from the Project tab of the ribbon. 

## Scenario

As a team member or project manager, I want quick access to view custom fields stored within a project. That means no desktop project management app and no navigation to reach the custom fields. 

### Using App

1.	Choose the PWA site you want to connect to and click connect.
2.  Upload the sample Project mpp file to PWA and publish
3.	Run the app (update project name if you are using a different project)

### Prerequisites
To use this code sample, you need the following:
* Project Online tenant.
* Visual Studio 2013 or later 
* Project CSOM client DLL.  It is available as a Nuget Package from [here](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM/)
* The project named "Local Custom Fields" added to the Project PWA site.
 
#### Deployment
1.	Load the "Local Custom Fields" project file onto the PWA site. 
2.	Update the user credentials in the app.
3.	Compile the app.


## How the sample affects your tenant data
This sample runs CSOM methods that reads the contents of uploaded project "Local Custom Fields" for the specified user. Tenant data will be affected.

## Additional resources
* [Local and Enterprise Custom Fields](https://msdn.microsoft.com/en-us/library/office/ms447495(v=office.14).aspx)

* [ProjectContext class](https://msdn.microsoft.com/en-us/library/office/microsoft.projectserver.client.projectcontext_di_pj14mref.aspx)

* [Client-side object model (CSOM)](https://aka.ms/project-csom-docs)

## Copyright

Copyright (c) 2016 Microsoft. All rights reserved.

