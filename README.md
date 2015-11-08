# PnP-Partner-Pack
Partner pack – Level set the partner ecosystem and customers on how to get started truly on the transformation and with typical SharePoint Add-In model implementations.
* Own repo as “PnP Partner Pack”
* Own version of the code, which uses PnP Core Nuget package
  * Only way we can ensure that code is stable enough and does compile
* Documentation and automated scripts for deployment
  * Step-by-Step <a href="./Documentation/Manual-Setup-Guide.md">guidance</a> on how to setup apps to Office 365 tenant and Azure
  * Architectural <a href="./DocumentationArchitecture-and-Implementation.md">documentation</a> about architecture and development

## Contents
* Site provisioning solution for site collections and sub sites, including “Save site as provisioning template” capability and centralized view.
* Responsive design template for sites, including custom navigation bar and footer with JavaScript embedding.
* Reference governance jobs implemented as Azure WebJobs.

## Objectives
* Easy Setup: Packaged solution with detailed <a href="./Documentation/Manual-Setup-Guide.md">guidance</a> on needed steps to configure solution to be running in Office 365 tenant. All you need is Office 365 tenant and Azure subscription.
* Ready to use reference solution: Starting point for your own solution with easy extension points, so that you do not need to start from scratch. Addresses most common customization scenarios, so that you can concentrate more on the business goals.
* Open Source: PnP Partner Pack is part of the Office 365 Developer Patterns and Practices (PnP) effort, which is open source community driven program.

## Value
* Each partner would have starting point for their customizations
* Each partner would have ready to use demo based on PnP guidance
* Level setting community with the model
* Concentration on sufficient guidance on getting started as easily as possible for learning purposes

### Why another solution for provisioning?
PnP Partner Pack is addressing specific scenario with simplified architecture.
There are few other solutions in PnP repository with different complexity levels, which we will continue also supporting in future.
