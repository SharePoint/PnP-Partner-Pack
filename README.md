# PnP Partner Pack 
This is the repository for PnP Partner Pack, which is part of the community driven [Office 365 Developer Patterns and Practices](http://aka.ms/OfficeDevPnP) (PnP) initiate. 

![](http://i.imgur.com/5L34MNk.png)

PnP Partner Pack can be considered as Starter Kit for customers and partners and combines numerous patterns and practices demonstrated in the [PnP samples](http://dev.office.com/patterns-and-practices-resources) to one reusable solution, which can be deployed and used in any Office 365 tenant. PnP Partner Pack is using [PnP SharePoint Core Component](https://github.com/OfficeDev/PnP-sites-core), which increases developer productivity with CSOM and REST based operations.

# What's included? #
PnP Partner Pack solution is demonstrating following capabilities:

- Self-service site collection and sub site provisioning solution
	- Fully configurable based on business requirements
	- Save existing site as new template from the standard user interface
	- Template creation does not require xml or script knowledge - New templates can be generated from the existing sites
	- Sub site creation implementation with remote provisioning
	- Support for tenant wide or site collection templates
- Responsive UI package for the Office 365 SP sites
	- Uses JavaScript and custom CSS files to transform oob SP sites as responsive
	- Can be applied to any SharePoint site and does not dependencies on the PnP Partner Pack
- UI widget implementations with JavaScript embedding pattern to avoid custom master pages
- Reference governance remote timer jobs (Azure WebJobs) to perform typical enterprise governance operations to existing site collections and sites
- Configurable branding and text elements for easy branding element changes

If you are interested on more detailed architectural description, please have a look on specific [PnP Partner Pack - Architecture documentation](./Documentation/Architecture-and-Implementation.md).

# How do I install PnP Partner Pack #
PnP Partner Pack requires quite a few installation steps. Here's resources around the installation and configuration.

- <a href="./Documentation/Manual-Setup-Guide.md">Step-by-Step guidance</a> on how to setup apps to Office 365 tenant and Azure
- [PnP Partner Pack - Setup Guide video](https://channel9.msdn.com/blogs/OfficeDevPnP/PnP-Partner-Pack-Setup-and-configuration) at Channel 9.


# What's the supportability story around this? #
Following statements apply cross all of the PnP samples and solutions, including PnP Partner Pack.

- PnP guidance and samples are created by Microsoft & by the Community
- PnP guidance and samples are maintained by Microsoft & community
- PnP uses supported and recommended techniques
- PnP implementations are reviewed and approved by Microsoft engineering
- PnP is open source initiative
- PnP is NOT a product and therefore it’s not supported through Premier Support or other official support channels
- PnP is supported in similar ways as other open source projects done by Microsoft with support from the community by the community
- There are numerous partners that utilize PnP within their solutions for customers. Support for this is provided by the Partner. When PnP material is used in deployments, we recommend to be clear with your customer / deployment owner on the support model.


# Can I use code and patterns from this solution? #
Yes. You can use this code and patterns anyway you want in your own implementations or offerings. You do not have to ask specific permission for anything around the PnP initiative. We rather hope that customers and partners would re-use our material as much as possible.

Obviously we are interested on your feedback and since PnP is open source community driven initiate, all contributions back to the PnP are absolutely welcome and appreciated, but not required.

# Screen shots of the solution #
Here's few screen shots of the different UIs solution provides for example for self-service site collection creation or for the sub site creation.

![](http://i.imgur.com/XAQgzVk.png)

Here's few screen shots of the responsive behavior included in the PnP Partner Pack.

![](http://i.imgur.com/y6iGZyk.png)

---

![](http://i.imgur.com/l01hhvE.png)