# PnP Partner Pack - Manual Setup Guide

##Solution Overview
PnP Partner pack allows you to extend the out of the box experience of Office 365 and 
SharePoint Online, by providing the following capabilities:
* Save Site as Provisioning Template feature in Site Settings
* Sub-Site creation  with custom UI and PnP Provisioning Template selection
* Site Collection creation for non-admin users with custom UI and PnP Provisioning Template selection
* My Site Collections personal view
* Responsive Design template for Site Collections
* Custom NavBar and Footer for Site Collections with JavaScript object model
* Sample Timer Jobs (implemented as WebJobs) for Governance rules enforcement

## Setup Overview
From a deployment perspective the PnP Partner Pack is an Office 365 Add-In, which 
leverages an Azure Web App with an Azure Web Sites and some Azure Web Jobs.

The application has to be registered in Azure Active Directory and acts against SharePoint
Online using an App Only access token, based on an X.509 self-signed Certificate.

Moreover, it is a requirement to have an Infrastructural Site Collection provisioned in
the target SharePoint Online tenant.

This document outlines the manual setup process, which allows you to leverage the 
PnP Partner Pack in your own environment.

## Manual Installation Steps
The manual installation requires to accomplish the following steps:
* [Azure Active Directory Application registration, as Office 365 Add-In](#azuread)
* [App Only certificate configuration in the Azure AD Application](#apponlyazuread)
* [Infrastructural Site Collection provisioning](#sitecollection)
* [Azure Web App provisioning and configuration](#azurewebapp)
* [App Only certificate configuration in the Azure Web App](#apponlywebapp)
* [Azure Web Jobs provisioning](#webjobs)

<a name="azuread"></a>
###Azure Active Directory Application Registration
First of all, because the PnP Partner Pack is an Office 365 Add-In, you have to register
it in the Azure Active Directory tenant that is under the cover of your Office 365 tenant.
In order to do that, open the Office 365 Admin Center (https://portal.office.com) using
the account of a user member of the Tenant Global Admins group.

Click on the "Azure AD" link that is available under the ADMIN group in the left-side 
treeview of the Office 365 Admin Center. In the new browser's tab that will be opened you
will find the Microsoft Azure Management Portal. If it is the first time that you access
the Azure Management Portal with your account, you will have to register a new Azure
subscription, providing some information and a credit card for any payment need.
But don't worry, in order to play with Azure AD and to register Office 365 Add-In you
will not pay anything. In fact, those are free capabilities. 
After having accessed the Azure Management Portal, select the "Active Directory" section,
by clicking on the icon highlighted in the following screenshot:

![Azure AD Button](./Figures/Fig-01-Azure-AD-Button.png)

On the right side of the screen you will se the Azure AD tenant corresponding to your
Office 365 tenant. Click on it to access its configuration, and then select the 
"Applications" tab. See the next figure for further details.

![Azure AD Main Page](./Figures/Fig-02-Azure-AD-Main-Page.png)

In the "Applications" tab you will find the list of Azure AD applications installed in 
your tenant. Click the "Add" button in the lower part of the screen, select the option
"Add an application my organization is developing".

![Azure AD - Add an Application - First Step](./Figures/Fig-03-Azure-AD-Add-Application-Step-01.png)

Then, provide a name for your application (we suggest to name it "OfficeDev PnP Partner
Pack"), and select the option "Web application and/or web API".

![Azure AD - Add an Application - Second Step](./Figures/Fig-04-Azure-AD-Add-Application-Step-02.png)

In the following registration step, provide the URL of the Azure Web App that you will
create (in one of the following steps) and define a unique ID, which can be for example
https://yourtenantonmicrosoft.com/OfficeDevPnP.PartnerPack.SiteProvisioning.

![Azure AD - Add an Application - Third Step](./Figures/Fig-05-Azure-AD-Add-Application-Step-03.png)

After having created the Azure AD Application, go into the "Configure" tab of the application.
There you can upload the app logo (<a href="https://raw.githubusercontent.com/OfficeDev/PnP-Partner-Pack/dev/OfficeDevPnP.PartnerPack.SiteProvisioning/OfficeDevPnP.PartnerPack.SiteProvisioning/PnP-O365-App-Icon.png">PnP-O365-App-Icon.png</a>),
and you can configure the application security.

First of all, you have to copy the value of the Client ID property. Moreover, you have
to configure a client secret. In order to do that, add a new security key (selecting 1 year
or 2 years for key duration). Press the "Save" button in the lower part of the screen to
generate the key value. After saving, you will see the key value. Copy it in a safe place,
because you will not see it anymore.

![Azure AD - Application Configuration - Client Secret](./Figures/Fig-06-Azure-AD-App-Config-01.png)

Scroll down a little bit more in the UI and configure the Application permissions, within
the "Permissions to other applications" section, which is illustrated in the following figure.

![Azure AD - Application Configuration - Client Secret](./Figures/Fig-07-Azure-AD-App-Config-02.png)

Click the "Add Application" button, a popup screen will appera. Select "Office 365 SharePoint Online"
and click the confirmation button. In the main configuration screen you have to configure
the following application permissions:

* Have full control of all site collection
* Read and write managed metadata

For further details, see the following figure.

![Azure AD - Application Configuration - Client Secret](./Figures/Fig-08-Azure-AD-App-Config-03.png)

The "Application Permissions" are those granted to the application when running as App Only.
The other dropdown of permissions, called "Delegated Permissions", defines the permissions
granted to the application when running under a specific user's account delegation (using
and app and user access token, from an OAuth 2.0 perspective). Because in the PnP Partner
Pack solution there are operations executed as App Only, and other execute through user's
delegation, you have to select the same permissions also for the "Delegated Permissions"
section. Then, click one more time the "Save" button.

<a name="apponlyazuread"></a>
###App Only certificate configuration in the Azure AD Application
You are now ready to configure the Azure AD Application for invoking SharePoint Online with
an App Only access token. In order to do that, you have to create and configure a self-signed
X.509 certificate, which will be used to authenticate your Application against Azure AD, while
requesting the App Only access token. 

First of all, you have to create the self-signer X.509 Certificate, which can be created 
using the makecert.exe tool that is available in the Windows SDK. If you have Microsoft
Visual Studio 2013/2015 installed on your enviroment, you already have the makecert tool, as well.
Otherwise, you will have to download from MSDN and to install the Windows SDK for your current
version of Windows Operating System.

The command for creating a new self-signed X.509 certificate is the following one:

```
makecert -r -pe -n "CN=MyCompanyName MyAppName Cert" -b 10/25/2015 -e 10/25/2016 -ss my -len 2048
```

The previous command creates a self-signed certificate with a common name (CN) value of "MyCompanyName MyAppName Cert", a validity
timeframe between 10/25/2015 and 10/25/2016, and a key length of 2048 bit. The certificate will have an exportable private key, 
and will be stored in the personal certificate store of the current user. 

>For further details about the makecert syntax and command line parameters you can read <a href="https://msdn.microsoft.com/en-us/library/windows/desktop/aa386968(v=vs.85).aspx">the following article</a>  on MSDN.

After having created the self-signed X.509 Certificate you have to export it as a .PFX file, which includes the private key value.
In order to do that, run the MMC.EXE command as an Administrator (RunAs Admin) and add the Certificates MMC snap-in, targeting the personal store of the current user.
In the Current User's Personal folder of Certificates, select the just created certificate, right click on it and select the "Export" functionality.
Select to export the private key into a .PFX file. Provide a password to protect the private key of the certificate.
Repeat the same process as before, but this time export the certificate as a .CER file, which does not include the private key value.

Start a PowerShell command window, and execute the following instructions:

```PowerShell
$certPath = Read-Host "Enter certificate path (.cer)"
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import($certPath)
$rawCert = $cert.GetRawCertData()
$base64Cert = [System.Convert]::ToBase64String($rawCert)
$rawCertHash = $cert.GetCertHash()
$base64CertHash = [System.Convert]::ToBase64String($rawCertHash)
$KeyId = [System.Guid]::NewGuid().ToString()

$keyCredentials = 
'"keyCredentials": [
    {
      "customKeyIdentifier": "'+ $base64CertHash + '",
      "keyId": "' + $KeyId + '",
      "type": "AsymmetricX509Cert",
      "usage": "Verify",
      "value":  "' + $base64Cert + '"
     }
  ],'
$keyCredentials

Write-Host "Certificate Thumbprint:" $cert.Thumbprint
```

Copy the output value into a text file, you will use it pretty soon.

Go back to the Azure AD Application that you created in the previous step and select the
"Manage Manifest" button in the lower area of the screen, then select the "Download Manifest" 
option in order to download the application manifest as a JSON file.

![Azure AD - Application Configuration - Client Secret](./Figures/Fig-09-Azure-AD-App-Config-04.png)

Open the just downloaded file using any text editor, search for the *keyCredentials* property and replace 
it with the following value.

```JSON
  "keyCredentials": [
    {
      "customKeyIdentifier": "<$base64CertHash>",
      "keyId": "<$KeyId>",
      "type": "AsymmetricX509Cert",
      "usage": "Verify",
      "value":  "<$base64Cert>"
     }
  ],
```

You can directly use the output of the previous PowerShell script, excluding the value of the certificate thumbprint.
Save the updated manifest and upload it back to Azure AD, by using the "Upload Manifest" functionality.

>For further details about running App Only applications, you can read <a href="http://blogs.msdn.com/b/richard_dizeregas_blog/archive/2015/05/03/performing-app-only-operations-on-sharepoint-online-through-azure-ad.aspx">the following article
>from Richard diZerega</a>.

<a name="sitecollection"></a>
###Infrastructural Site Collection provisioning

<a name="azurewebapp"></a>
###Azure Web App provisioning and configuration

<a name="apponlywebapp"></a>
###App Only certificate configuration in the Azure Web App

<a name="webjobs"></a>
###Azure Web Jobs provisioning
