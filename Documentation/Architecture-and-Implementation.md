# PnP Partner Pack - Architecture and Implementation Details

## Solution Overview
PnP Partner pack allows you to extend the out of the box experience of Microsoft Office 365 and 
Microsoft SharePoint Online, by providing the following capabilities:
* Save Site as Provisioning Template feature in Site Settings
* Sub-Site creation  with custom UI and PnP Provisioning Template selection
* Site Collection creation for non-admin users with custom UI and PnP Provisioning Template selection
* My Site Collections personal view
* Responsive Design template for Site Collections
* Governance tools for administrators: apply SharePoint farm-wide branding, refresh site templates, bulk creation of site collections 
* Custom NavBar and Footer for Site Collections with JavaScript Object Model
* Sample Timer Jobs (implemented as WebJobs) for Governance rules enforcement

In this document you will learn about the architecture and the implementation details of the
PnP Partner Pack.

For your convenience, here is the outline of the current document:
* [Architecture Overview](#architectureOverview)
* [Infrastructural Site Collection](#infrastructuralSiteCollection)
* [Azure Web App Details](#azureWebApp)
* [Infrastructure Library](#infrastructureLibrary)
* [Provisioning  Repository](#provisioningRepository)
* [Configuration Section](#configurationSection)
* [Asynchronous Jobs Handling](#asyncJobsHandling)
* [Creating Custom Jobs](#customJobs)
* [Important Notes](#importantNotes)

<a name="architectureOverview"></a>
## Architecture Overview
The overall architecture of PnP Partner Pack is based on an Azure Web App, which is
called *OfficeDev.PnPPartnerPack.SiteProvisioning*, and on Azure Active Directory.
In fact all the capabilities of the PnP Partner Pack are provided by an Azure 
AD Application, which is an Office 365 App, and which is hosted in Azure.
We decided to create the PnP Partner Pack as an Azure AD Application, instead of just
creating a SharePoint Add-In, because through the Azure AD Application we are
able to apply the extensions and to play with the app targeting whatever
SharePoint Online Site Collection, without the need to install and trust an Add-In
for each and every target Site Collection.

Moreover, the application runs against SharePoint
Online by leveraging an AppOnly access token, to allow to any user to leverage the provided capabilities.
The application also consumes the Microsoft Graph API with an OAuth access token related to the current
user identity. 

The application requires to have tenant-level permissions for the application by itself (AppOnly).
However, we assume that the requirement of having tenant-level permissions is not a road
blocking one, considering the target audience of the PnP Partner Pack.

Furthermore, there are a bunch of Web Jobs, which are provisioned within the same 
Azure Web App that hosts the main PnP Partner Pack Sites Provisioning Module.
From an implementation perspective, the Sites Provisioning module is implemented as an
ASP.NET MVC Web Application using Microsoft Visual Studio 2015/2017, and references the NuGet
package of the [OfficeDev PnP Core](#https://www.github.com/OfficeDev/PnP-Sites-Core/) library,
as well as all the related requirements.

<a name="infrastructuralSiteCollection"></a>
## Infrastructural Site Collection
At the very basic of the PnP Partner Pack infrastructure there is an Infrastructural Site Collection.
This Site Collection is created during the setup phase of the PnP Partner Pack, and contains some 
fundamental artifacts like:
* *PnPProvisioningJobs*: Document Library that stores Provisioning Jobs. You can learn more about this
topic in the [Asynchronous Jobs Handling](#asyncJobsHandling) section.
* *PnPProvisioningTemplates*: Document Library that stores tenant-level global Provisioning Templates. It
is used to share PnP Provisioning Templates (manually defined or saved from an existing site) that can
be applied while provisioning sites across all the Site Collections of the current tenant.

Having a PnP Partner Pack Infrastructural Site Collection is a mandatory requirement for the solution.

From a Provisioning Templates perspective, it is also possible to store them locally at the 
Site Collection level. In that case, a *PnPProvisioningTemplates* library will be provisioned in the
target Site Collection. Templates stored locally in a Site Collection will be available only when creating
sites in that specific Site Collection. Furthermore, local templates are available only
for Sub-Sites creation.

<a name="azureWebApp"></a>
## Azure App Service Details
As already stated, the Sites Provisioning Web application is an ASP.NET MVC Web Application
that basically provides the implementation for three controllers:
* *HomeController*: implements the main actions of the PnP Partner Pack, in order to manage sites creation and templates management.
* *PersonaController*: provides basic functionalities to manage persona controls, users' profile picture, etc.
* *GovernanceController*: implements the main governance actions of the PnP Partner Pack.

The *HomeController* provides the logic for handling the following capabilities:
* Save Site As Provisioning Template: based on the *SaveSiteAsTemplate* action.
* Site Collection Provisioning UI: based on the *CreateSiteCollection* action. 
* Sub-Site Provisioning UI: based on the *CreateSubSite* action.
* Settings capability: based on the *Settings* action.
* My Site Collections UI: based on the *MyProvisionedSites* action.

The *GovernanceController* provides the logic for handling the following capabilities:
* Apply SharePoint-wide branding to all sites and site collections in the Office 365 tenant.
* Refresh templates applied to sites and site collections created using the PnP Partner Pack v. 2.0 
* Batch creation of site collections using a bulk loaded XML file

Of course, each main action has a corresponding View and a Model, from a MVC perspective.
The only actions that deserve a specific explanation are those related to Site Collections and
Sub-Sites creation. In fact, we decided to leverage some OOP techniques and a wizard-like UI/UX.
Both the *CreateSiteCollection* and the *CreateSubSite* actions are related to corresponding
*CreateSiteCollectionViewModel* and *CreateSubSiteViewModel* Models, which inherit from a common
abstract base type called *CreateSiteViewModel*. From a View perspective, there is a unique View
defined in the CreateSite.cshtml file, which defines the global wrapper of the Site/Sub-Site 
provisioning wizard.
Then, the CreateSite.cshtml view leverages a set of partial views to define the steps of the wizard.

The Site Collection provisioning wizard leverages the following partial views:
* TemplateSelection.cshtml 
* SiteCollectionInformation.cshtml
* TemplateParameters.cshtml
* SiteCreated.cshtml

While the Sub-Site provisioning wizard leverages the following partial views:
* TemplateSelection.cshtml 
* SubSiteInformation.cshtml
* TemplateParameters.cshtml
* SiteCreated.cshtml

As you can see the only difference is the View related to collecting information about the 
provisioning target. In fact, provisioning a Site Collection requires some information that are
different from those required to provision a Sub-Site. However, all the other wizard steps will share
the same base Model (*CreateSiteViewModel*) to store generic provisioning configurations.

We used a similar approach for the Batch creation of Site Collections in the Governance section. In fact, there we use the following views:
* BatchStartup.cshtml
* BatchFileUploaded.cshtml
* BatchScheduled.cshtml

In order to consume the Microsoft Graph API, the application uses Active Directory Authentication Library (ADAL).
Notice that ADAL uses a token cache for OAuth tokens. For the sake of simplicity, the token cache provided is
based on the web application session and is implemented in type *SessionTokenCache*. 
However,  a session-based token cache is not a scalable solution and it cannot be used with multiple instances of the web app.
Nevertheless, you can configure a session based on an external persistence provider, like for example the
<a href="https://azure.microsoft.com/en-us/documentation/articles/cache-asp.net-session-state-provider/">Azure Redis Cache</a>,
or you can define a token cache handler of your own, using a backend database or
whatever else. For further details about ADAL and the token cache, you can read the book
<a href="https://www.microsoftpressstore.com/store/modern-authentication-with-azure-active-directory-for-9780735696945">"Modern Authentication with Azure
Active Directory for Web Applications"</a> written by <a href="http://www.cloudidentity.com/">Vittorio Bertocci</a>.

Aside from that, the Azure App Service is a very common ASP.NET MVC Web Application, which internally 
leverages the PnP Partner Pack Infrastructural Library to accomplish any real business task.

<a name="infrastructureLibrary"></a>
## Infrastructure Library
This is a .NET class library, which is defined in the project named *OfficeDevPnP.PartnerPack.Infrastructure*,
that provides all the core functionalities of the PnP Partner Pack.
In this library you will find the following main types:
* *IConfigurable*: defines the common interface that any configurable type should implement.
* *IProvisioningRepository*: declares the basic interface for any concrete Provisioning Repository.
It allows to define an abstract persistence layer. By default the PnP Partner Pack provides a 
Provisioning Repository that is backed on SharePoint Online, using the 
[Infrastructural Site Collection](#infrastructuralSiteCollection).
* *ITemplatesProvider*: defines the common interface for any concrete Templates Provider service, 
in order to being able to leverage external template providers like the PnP Templates Gallery.
* *PnPPartnerPackConstants*: declares a bunch of constant values used around the solution.
* *PnPPartnerPackContextProvider*: this is a fundamental type that handles the creation of a CSOM
*ClientContext* object, bases on the current configuration. It hides the complexity of creating AppOnly
*ClientContext* instances.
* *PnPPartnerPackSettings*: provides a quick and direct path to access all the settings related to the
PnP Partner Pack in the current context.
* *PnPPartnerPackUtilities*: provides a set of useful helper methods to play with the capabilities 
offered by the PnP Partner Pack. For example here you will find methods to apply a Provisioning Template
to a target site, or to enable or disable the PnP Partner Pack extensions onto a target site, etc.
* *ProvisioningRepositoryFactory*: it is a factory class that allows creating a concrete instance of the
currently configured Provisioning Repository type.
* *SharePointProvisioningRepository*: is defined in the namespace *OfficeDevPnP.PartnerPack.Infrastructure.SharePoint*
and is the out of the box available implementation of a Provisioning Repository, which targets SharePoint
Online and the [Infrastructural Site Collection](#infrastructuralSiteCollection).

<a name="provisioningRepository"></a>
## Provisioning  Repository
The Provisioning Repository is a .NET type that manages the persistence storage layer. Any Provisioning
Repository has to implement the *IProvisioningRepository* interface, which is defined like the following
code excerpt.

```csharp
public interface IProvisioningRepository : IConfigurable
{
    /// <summary>
    /// Retrieves the list of Global Provisioning Templates
    /// </summary>
    /// <param name="scope">The scope to filter the provisioning templates</param>
    /// <returns>Returns the list of Provisioning Templates</returns>
    ProvisioningTemplateInformation[] GetGlobalProvisioningTemplates(TemplateScope scope);

    /// <summary>
    /// Retrieves the list of Local Provisioning Templates
    /// </summary>
    /// <param name="siteUrl">The local Site Collection to retrieve the templates from</param>
    /// <param name="scope">The scope to filter the provisioning templates</param>
    /// <returns>Returns the list of Provisioning Templates</returns>
    ProvisioningTemplateInformation[] GetLocalProvisioningTemplates(String siteUrl, TemplateScope scope);

    /// <summary>
    /// Saves a Provisioning Template into the target Global repository
    /// </summary>
    /// <param name="template">The Provisioning Template to save</param>
    void SaveGlobalProvisioningTemplate(GetProvisioningTemplateJob job);

    /// <summary>
    /// Saves a Provisioning Template into the target Local repository
    /// </summary>
    /// <param name="siteUrl">The local Site Collection to save to</param>
    /// <param name="template">The Provisioning Template to save</param>
    void SaveLocalProvisioningTemplate(String siteUrl, GetProvisioningTemplateJob job);

    /// <summary>
    /// Enqueues a new Provisioning Job
    /// </summary>
    /// <param name="job">The Provisioning Job to enqueue</param>
    /// <returns>Returns the ID of the job</returns>
    Guid EnqueueProvisioningJob(ProvisioningJob job);

    /// <summary>
    /// Updates a job in the queue
    /// </summary>
    /// <remarks>In case of failure it will throw an Exception</remarks>
    /// <param name="job">The information about the job to update</param>
    void UpdateProvisioningJob(Guid jobId, ProvisioningJobStatus status, String errorMessage = null);

    /// <summary>
    /// Retrieves the list of Provisioning Jobs
    /// </summary>
    /// <param name="status">The status to use for filtering Provisioning Jobs</param>
    /// <param name="includeStream">Defines whether to include the stream of the serialized job</param>
    /// <param name="owner">The optional owner of the Provisioning Job</param>
    /// <returns>The list of information about the Provisioning Jobs, if any</returns>
    ProvisioningJobInformation[] GetProvisioningJobs(ProvisioningJobStatus status, String jobType = null, Boolean includeStream = false, String owner = null);

    /// <summary>
    /// Retrieves a Provisioning Job by ID
    /// </summary>
    /// <param name="jobId">The ID of the job to retrieve</param>
    /// <param name="includeStream">Defines whether to include the stream of the serialized job</param>
    /// <returns>The information about the Provisioning Job, if any</returns>
    ProvisioningJobInformation GetProvisioningJob(Guid jobId, Boolean includeStream = false);

    /// <summary>
    /// Retrieves the list of Provisioning Jobs
    /// </summary>
    /// <param name="status">The status to use for filtering Provisioning Jobs</param>
    /// <param name="owner">The optional owner of the Provisioning Job</param>
    /// <typeparam name="TJob">Represents the type of the Provisioning Jobs to retrieve</typeparam>
    /// <returns>The list of information about the Provisioning Jobs, if any</returns>
    ProvisioningJob[] GetTypedProvisioningJobs<TJob>(ProvisioningJobStatus status, String owner = null)
        where TJob : ProvisioningJob;
}
```

As you can see the *IProvisioningRepository* interface completely decouples the PnP Partner Pack from
any provisioning target. As already stated, the engine out of the box provides the
*SharePointProvisioningRepository* concrete implementation, but you are free to create your own.
For example, you could create a Provisioning Repository that targets Azure SQL Database and/or
Azure DocumentDB. In that case, please feel free also to share your own implementation with the
whole community by submitting a Pull Request into the [dev branch of this GitHub repository](../tree/dev). 

<a name="configurationSection"></a>
## Configuration Section
The PnP Partner Pack leverages a custom configuration section, which allows to define some custom 
settings related to the current Office 365 tenant, the Azure AD Application settings, as well as the
configuration of the Provisioning Jobs.

Here follows a sample excerpt of the XML configuration section for the PnP Partner Pack.

```XML
  <!-- PnP Partner Pack Settings -->
  <PnPPartnerPackConfiguration xmlns="http://schemas.dev.office.com/PnP/2015/10/PnPPartnerPackConfiguration">
    <GeneralSettings defaultSiteTemplate="STS#0"
                     Title="PnP Partner Pack"
                     LogoUrl="/AppIcon.png">
      <WelcomeMessage>
        <![CDATA[
          Welcome to the PnP Partner Pack, which is a project managed by the <a href="http://aka.ms/OfficeDevPnP" target="_blank">Office 365 Developers Patterns &amp; Practices</a> team!<br />
          This is a sample solution, including source code, that illustrates to the partners' ecosystem and customers how to get started truly on the transformation, and with typical SP add-in model implementations.<br />
          Here you can find samples about how to manage the provisioning of Site Collection or Sub Sites, applying one or more provisioning templates.<br />
          The provisioning is based on the new Remote Provisioning technique, by leveraging the PnP Provisioning Engine.<br />
          Let's play with this sample solution and enjoy the new Add-In Model for Microsoft SharePoint andd Microsoft Office 365.
        ]]>
      </WelcomeMessage>
      <FooterMessage>
        <![CDATA[
          <p>
            &copy; <a href="http://aka.ms/OfficeDevPnP">Office 365 Developers Patterns &amp; Practices</a>
          </p>
        ]]>
      </FooterMessage>
    </GeneralSettings>

    <TenantSettings tenant="[tenant].onmicrosoft.com" 
      appOnlyCertificateThumbprint="[X.509 Self-Signed Certificate Thumbprint]" 
      infrastructureSiteUrl="https://[tenant].sharepoint.com/sites/PnP-Partner-Pack-Infrastructure/" />

    <ProvisioningRepository name="SharePointProvisioningRepository" 
      type="OfficeDevPnP.PartnerPack.Infrastructure.SharePoint.SharePointProvisioningRepository, OfficeDevPnP.PartnerPack.Infrastructure" />

    <TemplatesProviders>
      <TemplatesProvider name="TenantGlobal" enabled="true" type="OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders.SharePointGlobalTemplatesProvider, OfficeDevPnP.PartnerPack.Infrastructure" />
      <TemplatesProvider name="SiteCollectionLocal" enabled="true" type="OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders.SharePointLocalTemplatesProvider, OfficeDevPnP.PartnerPack.Infrastructure" />
      <TemplatesProvider name="TemplatesGallery" enabled="true" type="OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders.PnPTemplatesGalleryProvider, OfficeDevPnP.PartnerPack.Infrastructure">
        <Configuration>
          <gallery url="https://templates-gallery.sharepointpnp.com/" />
        </Configuration>
      </TemplatesProvider>
    </TemplatesProviders>

    <ProvisioningJobs>
      <JobHandlers>
        <JobHandler name="ProvisioningTemplateJobHandler" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers.ProvisioningTemplateJobHandler, OfficeDevPnP.PartnerPack.Infrastructure" />
        <JobHandler name="SiteCollectionProvisioningJobHandler" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers.SiteCollectionProvisioningJobHandler, OfficeDevPnP.PartnerPack.Infrastructure" />
        <JobHandler name="SubSiteProvisioningJobHandler" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers.SubSiteProvisioningJobHandler, OfficeDevPnP.PartnerPack.Infrastructure" />
        <JobHandler name="BrandingJobHandler" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers.BrandingJobHandler, OfficeDevPnP.PartnerPack.Infrastructure" />
        <JobHandler name="SiteCollectionsBatchJobHandler" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers.SiteCollectionsBatchJobHandler, OfficeDevPnP.PartnerPack.Infrastructure" />
        <JobHandler name="RefreshSitesJobHandler" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers.RefreshSitesJobHandler, OfficeDevPnP.PartnerPack.Infrastructure" />
      </JobHandlers>
      <JobTypes>
        <JobType handler="ProvisioningTemplateJobHandler" executionModel="Scheduled" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.GetProvisioningTemplateJob" />
        <JobType handler="ProvisioningTemplateJobHandler" executionModel="Continous" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.ApplyProvisioningTemplateJob" />
        <JobType handler="SiteCollectionProvisioningJobHandler" executionModel="Scheduled" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.SiteCollectionProvisioningJob" />
        <JobType handler="SubSiteProvisioningJobHandler" executionModel="Continous" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.SubSiteProvisioningJob" />
        <JobType handler="BrandingJobHandler" executionModel="Scheduled" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.BrandingJob" />
        <JobType handler="SiteCollectionsBatchJobHandler" executionModel="Scheduled" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.SiteCollectionsBatchJob" />
        <JobType handler="RefreshSitesJobHandler" executionModel="Scheduled" type="OfficeDevPnP.PartnerPack.Infrastructure.Jobs.RefreshSitesJob" />
      </JobTypes>
    </ProvisioningJobs>

  </PnPPartnerPackConfiguration>
```

As you can see the main configuration elements are:
* *GeneralSettings*: defines the default site template to use when provisioning a new site, and before
applying the PnP Provisioning Template. We suggest using the STS#0 (SharePoint "Team Site"), which is
the most commonly used and the one that we tested more. Of course, you are free to change this option.
It also defines the LogoUrl and the Title that will be used to rendere the pages of the PnP Partner Pack.
Moreover, it defines the Welcome Message and the Footer Message for the Home Page and the global layout
view of the MVC site that renders the Site Provisioning engine.
* *TenantSettings*: defines the information about the target Office 365 tenant, the thumbprint of the
X.509 certificate to use for AppOnly authentication against Azure AD and the target SharePoint Online, 
and the URL of the Infrastructural Site Collection.
* *ProvisioningRepository*: defines the concrete type to use as the Provisioning Repository. You should
change this configuration in order to use a custom Provisioning Repository.
* *TemplatesProviders*: defines the list of templates providers that can be used to retrieve PnP 
provisioning templates when you have to create a new site or site collection. Out of the box the PnP
Partner Pack supports SharePoint Online (tenant level and local site collection level), as well as the
PnP Templates Gallery public and open source repository.
* *ProvisioningJobs*: declares what are the Job Handlers, the Job Types, and the execution model of the
Job Types. An execution model value of *Scheduled* means that the Job will be executed by the
*ScheduledJob*, while a value of *Continous* means that the Job will be executed by the *ContinousJob*.
If you create any custom Job, you will have to define the Job Type and the Job  Handler through this
elements. For further details about the Job Handling, please read the following section.

<a name="asyncJobsHandling"></a>
## Asynchronous Jobs Handling
The Site and Sub-Site provisioning actions simply provide the UI to collect data from the
end user, but do not really provision the corresponding target. On the contrary, whenever
you fill out the forms for feeding the provisioning of a Site or a Sub-Site, an asynchronous
job will be created. In the background there are two Web Jobs that handle the requests
based on a list of enqueued jobs.

These two Web Jobs are:
* *OfficeDevPnP.PartnerPack.ContinousJob*: handles jobs using a near to real-time approach. 
Whenever a job is created, the engine also puts a message in an Azure Blob Storage Queue. 
That message will wake up the *ContinousJob* and make the near to real-time behavior possible.
* *OfficeDevPnP.PartnerPack.ScheduledJob*: handles jobs based on a configurable schedule 
(every 5 minutes, every 10 minutes, every hour, or whatever else).

More in general, the jobs are stored in a Document Library (called *PnPProvisioningJobs*), 
which is provisioned in the Infrastructural Site Colleciton of the PnP Partner Pack during the 
setup phase. Every single job item is a document (that is a JSON file) that represents the 
JSON serialization of an entity that is defined in the *OfficeDevPnP.PartnerPack.Infrastructure* project
and that inherits from the *ProvisioningJob* type, which is defined in the 
*OfficeDevPnP.PartnerPack.Infrastructure.Jobs* namespace.

Out of the box the PnP Partner Pack defines the following types of jobs:
* *SiteCollectionProvisioningJob*: provisions a new Site Collection.
* *SubSiteProvisioningJob*: provisions a new Sub-Site.
* *GetProvisioningTemplateJob*: extracts the PnP Provisioning Template from an existing source.
* *ApplyProvisioningTemplateJob*: applies a PnP Provisioning Template to a target. 
* *BrandingJob*: applies the SharePoint-wide branding settings to all sites and site collections.
* *SiteCollectionsBatchJob*: handles the batch creation of a bunch of site collections.
* *RefreshSitesJob*: refreshes the provisioning template of all the sites and site collections created with the PnP Partner Pack v. 2.0.

Every single Job document is stored in the *PnPProvisioningJobs* Document Library with some
custom metadata fields, which allow to define the type of job (can be a Site Collection Provisioning,
a Sub-Site provisioning, a Site Collection template extraction or application, etc). There is also a
metadata field that declares the status of the job. The available values for the job status are:
* *Pending*: the job is waiting to be handled.
* *Failed*: the job handling failed, you will find detailed information about the issue in 
the "Provisioning Job Error" metadata field.
* *Cancelled*: the job has been cancelled by the  end user.
* *Running*: the job is being handled right now.
* *Provisioned*: the job has already been applied and provisioned.

Every single job type is handled by a Job Handler, which again is a custom type that inherits
from the *ProvisioningJobHandler* type declared in the 
*OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers* namespace.
Here are the out of the box available Job Handlers:
* *SiteCollectionProvisioningJobHandler*: handles provisioning of Site Collections
(i.e. *SiteCollectionProvisioningJob* instances).
* *SubSiteProvisioningJobHandler*: handles provisioning of Sub-Sites
(i.e. *SubSiteProvisioningJob* instances).
* *ProvisioningTemplateJobHandler*: handles extraction and application of PnP Provisioning Templates
(i.e. *GetProvisioningTemplateJob* and *ApplyProvisioningTemplateJob* instances).
* *BrandingJobHandler*: handles the branding jobs (i.e. BrandingJob)
* *SiteCollectionsBatchJobHandler*: handles the batch creation of site collections (i.e. SiteCollectionsBatchJob)
* *RefreshSitesJobHandler*: handles the refresh of templates for sites and site collections created with the PnP Partner Pack v. 2.0 (i.e. RefreshSitesJob)

Based on the kind of job and on some configuration elements in the .config files of the PnP Partner Pack
modules, the jobs can be executed by the *ContinousJob*, or by the *ScheduledJob*. In fact,
if you look into the .config file of the Sites Provisioning application (or into any of the .config files
of the two Provisioning Web Jobs) you will find a configuration section element called *&lt;ProvisioningJobs /&gt;*.

> For further details about the configuration of PnP Partner Pack, read the section
["Configuration Section"](#configurationSection).

<a name="customJobs"></a>
### Creating Custom Jobs
If you want, you can create your own *ProvisioningJob* types, simply by inheriting from the
base abstract class called *OfficeDevPnP.PartnerPack.Infrastructure.Jobs.ProvisioningJob*.
Moreover, you will need to define a custom Job Handler, which can be created by inheriting from
the base abstract class *ProvisioningJobHandler*. You will have to implement the *RunJobInternal* 
abstract method, and you will have also to implement the *Init* abstract method.
Just after that, you will have to configure your new job and its handler in the .config file of the
target PnP Partner Pack Environment.
Of course, you will also have to determine if you want to run your custom job with the *ContinousJob* or 
by using the *ScheduledJob*. This choice will determine the configuration that you will have to write
in the .config file.
Just after that, you will be ready to enqueue your custom jobs using the *EnqueueProvisioningJob* method
of the currently configured Provisioning Repository.

> For further details about the Provisioning Repository in the PnP Partner Pack, read the section
["Provisioning Repository"](#provisioningRepository).

<a name="importantNotes"></a>
## Important Notes
Here is a list of important things to know in order to master the PnP Partner Pack solution.

### AppOnly Access Token and Authorization Rules
Keep in mind that the PnP Partner Pack runs with an AppOnly token against SharePoint Online. This
allows any user to play with the solution. However, the current implementation of the solution does
not provide any authorization rules. We advise you to customize the solution, which is open source,
in order to include your own custom authorization rules.

### Taxonomies Support
Because the the PnP Partner Pack runs with an AppOnly token against SharePoint Online, it does not 
support applying taxonomies while applying any PnP Provisioning Template. Thus, if you are going to 
apply a PnP Provisioning Template that includes Term Groups, those will be remove from the template
before applying it.

### Managed Navigation Support
Because the the PnP Partner Pack runs with an AppOnly token against SharePoint Online, it does not 
support applying managed navigation for sites. If you apply a PnP Provisioning Template, which includes
managed navigation settings (current or global), the PnP Partner Pack will remove that navigation settings 
and will not apply it onto the target site.

### Search Settings Support
Because the the PnP Partner Pack runs with an AppOnly token against SharePoint Online, it does not 
support applying search settings while applying any PnP Provisioning Template. Thus, if you are going to 
apply a PnP Provisioning Template that includes Search Settings, those will be remove from the template
before applying it.

### Telemetry
In order to measure and track the usage of the PnP Partner Pack, we introduced a "call home" function in the Index action
of the Home controller of the PnP Partner Pack. Thus, we will be able to monitor and track what are the Office 365 tenants 
that benefit from using the PnP Partner Pack. Moreover, every single view (.CSHTML) of the project includes a 1 pixel tracking
image tag, at the very end of the view, in order to track usage of the various functionalities of the PnP Partner Pack.
This kind of monitoring and telemetry allows us to better invest our time and money in implementing and improving what people
in the community really use.
However, if you don't like to have tracking and telemetry in your deployment, you can remove the "call home" function in the
Index action of the Home controller, and you can remove the image elements at the end of all of the view. This is an open source
project, thus you have the source code and you can do whatever you like with it.

<img src="https://telemetry.sharepointpnp.com/pnp-partner-pack/documentation/architecture-and-implementation" /> 

