# Task module and Deep linking

## Summary

This sample is an enhancement of the personal tab created from using the core teams application built over the course of labs [A01](../../../../lab-instructions/aad/A01-begin-app.md)-[A03](../../../../lab-instructions/aad/A03-after-apply-styling.md).

In this sample, users can open a dialog(task module) from within the personal tab to view north wind orders' sales team information.

<img src="../../../../assets/07-005-chat.png?raw=true" alt="chat"/>

And in the dialog, a button to initiate a group chat using deep linking to discuss a particular order with the sale team.

<img src="../../../../assets/07-005-chat01.png?raw=true" alt="chat"/>

## Version history

Version|Date|Author|Comments
-------|----|----|--------
1.0|February 2022|Rabia Williams|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone or download the sample from [https://github.com/OfficeDev/m365-msteams-northwind-app-samples](https://github.com/OfficeDev/m365-msteams-northwind-app-samples)

- In a console, navigate to `src/extend-with-capabilities/aad/DeeplinksTaskModules/` from the main folder `m365-msteams-northwind-app-samples`.

    ```bash
    cd src/extend-with-capabilities/aad/DeeplinksTaskModules/
    ```

- Install modules

    ```bash
    npm install
    ```

- Run ngrok - point to port 3978

    ```bash
    ngrok http -host-header=rewrite 3978
    ```

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

This is a simple teams personal tab application, which shows how you can open a dialog (task module) in a tab to display information and use deep linking to open a group chat within the same dialog.

## Debug and test locally

TBD