# Microsoft Bot Framework Webchat - Sharepoint Web Part

## Summary

A web part that acts as a web chat component for bot's built on the Microsoft Bot Framework using the DirectLine API. When sending messages the web part uses the username of the currently logged in user.

## Used SharePoint Framework Version

![1,11.0](https://img.shields.io/badge/drop-1.11.0-green.svg)

## Applies to

* [SharePoint Framework Developer](https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
* [Office 365 developer tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
* [Microsoft Bot Framework](http://dev.botframework.com)

## Prerequisites

You need to have a bot created and registered using the Microsoft Bot Framework and registered to use the DirectLine Channel,
which will give you the secret needed inorder to generate the token when adding this web part to the page.  For more information on creating a bot and registering the channel you can see the official web site at [dev.botframework.com](http://dev.botframework.com).

## Building the code
- Clone this repository
- Update the bot token API endpoint & ResourceId in the manifiest file (BotWebPart.manifest.json)
- in the command line run:
  - `npm install`
  - `npm install -g gulp`
  - `gulp serve`

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* sharepoint/* - all resources which should be uploaded to a CDN.

## Connect Securely
Update the Web part properties to choose one of the following secured bot connectivity options
- Direct Line Secret
  - Copy the direct line secret from the registered direct line channel. 
- Direct Line Token API
  - Use the bot direct line token API to get the token to securely connect the bot.
- Custom API Secured by OAuth
  - Use your own custom API secured by OAuth which inturn generates the bot direct line token using the direct line token API.
  - To leverage this connectivity option, please update the following properties in the "BotWebPart.manifest.json" file
    - botTokenApiResourceId
    - botTokenApiUrl  

Direct Line Token API and Custom API is the recommended one to more securely connect your bot.

## Package Options

* gulp clean
* gulp build
* gulp serve
* gulp bundle --ship
* gulp package-solution --ship

## Features
This Web Part illustrates the following concepts on top of the SharePoint Framework:

- Connecting and communicating with a bot built on the Microsoft Bot Framework using the DirectLine Channel
- Office UI Fabric
- React

# Project

This repo has been populated by an initial template to help get you started. Please make sure to update the content to build a great experience for community-building.

As the maintainer of this project, please make a few updates:

- Improving this README.MD file to provide a great experience
- Updating SUPPORT.MD with content about this project's support experience
- Understanding the security reporting process in SECURITY.MD
- Remove this section from the README

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.

