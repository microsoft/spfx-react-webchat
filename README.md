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
