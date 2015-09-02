# OfficeFabricUIOffice365App

## Overview ##
This is a sample project of an Office 365 application implementing Office Fabric UI. This application is a standard MVC application created with Visual Studio 2013, using [Office 365 APIs client libraries](http://aka.ms/kbwa5c) with the new [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric).

## Prerequisites and Configuration ##

Following prerequisites and configuration are similar than existing samples provided on [OfficeDev GitHub Repo](https://github.com/OfficeDev/) like the [Office 365 Start Project for ASP.NET MVC](https://github.com/OfficeDev/O365-ASPNETMVC-Start/)

This sample requires the following:

  - Visual Studio 2013 with Update 3 or Visual Studio 2015.
  - [Microsoft Office 365 API Tools version 1.4.50428.2](http://aka.ms/k0534n). 
  - An [Office 365 developer site](http://aka.ms/ro9c62) or another Office 365 tenant.
  - Microsoft IIS enabled on your computer.

### Register app and configure the sample to consume Office 365 APIs ###

You can do this via the Office 365 API Tools for Visual Studio (which automates the registration process). Be sure to download and install the [Office 365 API tools](http://aka.ms/k0534n) from the Visual Studio Gallery before you proceed any further.

   1. Build the project. This will restore the NuGet packages for this solution. 
   2. In the Solution Explorer window, choose **O365-APIs-Start-ASPNET-MVC** project -> **Add** -> **Connected Service**.
   2. A Services Manager window will appear. Choose **Office 365** -> **Office 365 APIs** and select the **Register your app** link.
   3. If you haven't signed in before, a sign-in dialog box will appear.  Enter the user name and password for your Office 365 tenant admin. We recommend that you use your Office 365 Developer Site. Often, this user name will follow the pattern {username}@{tenant}.onmicrosoft.com. If you do not have a developer site, you can get a free Developer Site as part of your MSDN Benefits or sign up for a free trial. Be aware that the user must be a Tenant Admin user—but for tenants created as part of an Office 365 Developer Site, this is likely to be the case already. Also developer accounts are usually limited to one user.
   4. After you're signed in, you will see a list of all the services. Initially, no permissions will be selected, as the app is not registered to consume any services yet. 
   5. To register for the services used in this sample, choose the following permissions, and select the Permissions link to set the following permissions:
	- (Contacts) – Read 
	- (Mail) - Read
	- (Users and Groups) – Sign you in and read your profile (Read)
   6. Choose the **App Properties** link in the Services Manager window. Make this app available to a Single Organization. 
   7. After selecting **OK** in the Services Manager window, assemblies for connecting to Office 365 REST APIs will be added to your project and the following entries will be added to your appSettings in the web.config: ClientId, ClientSecret, AADInstance, and TenantId. You can use your tenant name for the value of the TenantId setting instead of using the tenant identifier.
   8. Build the solution. Nuget packages will be added to you project. Now you are ready to run the solution and sign in with your organizational account to Office 365.

