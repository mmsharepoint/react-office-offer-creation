# Offer Creation (SPFx) - Microsoft Teams App

## Summary

This sample is a Teams personal Tab to act as a Microsoft 365 across application (Teams, Outlook, Office) including a search-based messaging extension to act in Teams and Outlook. It is realized with SharePoint Framework (SPFx).

App live in action inside Teams

![App live in action inside Teams](assets/16OfferCreationDemo_SPFx.gif)

Create Offer form with FluentUI controls

![Create Offer form with FluentUI controls](assets/15CreateOfferForm_FluentUI_SPFx.png)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.16.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?


## Version history

Version|Date|Author|Comments
-------|----|----|--------
1.0|Dec xx, 2022|[Markus Moeller](https://twitter.com/moeller2_0)|Initial release

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

- Create the content-type for your offers in a site / default document library of your choice
    - With PnP-PowerShell for instance call the deploy script with your site url as parameter
        ```bash
        .\templates\deploy.ps1 -siteUrl <YourFullSiteUrl>
    
    - This should put the same site url to your tenant-property named 'CreateOfferSiteUrl'


## Features

* Using SharePoint Rest API to copy files and edit it's metadata
* [Extend Teams SPFx apps across Microsoft 365](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/office/overview?WT.mc_id=M365-MVP-5004617)
* [Use FluentUI Label, DatePicker, Dropdown, IDropdownOption, Spinner, TextField](https://developer.microsoft.com/en-us/fluentui#/?WT.mc_id=M365-MVP-5004617)
* [Use SharePoint tenant properties for org-wide SPFx app configurations](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties?tabs=sprest#getread-tenant-properties?WT.mc_id=M365-MVP-5004617)

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
