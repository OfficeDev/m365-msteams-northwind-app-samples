# Extend Teams app to other M365 host apps like Outlook, Office.com


## Summary

This sample is a created  using the core teams application built over the course of labs [A01](../../../../lab-instructions/aad/A01-begin-app.md)-[A03](../../../../lab-instructions/aad/A03-after-apply-styling.md) to get to the Northwind Orders core application. The app demonstrates how to use the latest Microsoft Teams JS SDK V2 to extend teams application to other M365 host apps like Outlook/Office.com

## Frameworks

![drop](https://img.shields.io/badge/Bot&nbsp;Framework-4.7-green.svg)


## Version history

Version|Date|Author|Comments
-------|----|----|--------
1.0|April 2022|Rabia Williams|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone or download the sample from [https://github.com/OfficeDev/m365-msteams-northwind-app-samples](https://github.com/OfficeDev/m365-msteams-northwind-app-samples)

- In a console, navigate to `src/extend-with-capabilities/ExtendTeamsForM365/` from the main folder `m365-msteams-northwind-app-samples`.

    ```bash
    cd src/extend-with-capabilities/ExtendTeamsForM365/
    ```

- Install modules

    ```bash
    npm install
    ```

- Run ngrok - point to port 3978

    ```bash
    ngrok http -host-header=rewrite 3978
    ```

- Update the `.env` configuration 


- Package the app

    ```bash
    npm run package
    ```

- Run the bot locally
    ```bash
    npm start
    ```

- Upload the the packaged zip file inside `manifest` folder into Teams [using these instructions](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-upload).



## Features



## Debug and test locally

TBD