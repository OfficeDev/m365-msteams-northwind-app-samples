# Search Messaging Extension

## Summary

This sample is a Search messaging extension created from using the core teams application built over the course of labs [A01](../../../../lab-instructions/aad/A01-begin-app.md)-[A03](../../../../lab-instructions/aad/A03-after-apply-styling.md) to get to the Northwind Orders core application It inserts information about Northwind customers into the compose box.

Users can search the Northwind database when composing a message and find the product.


And insert an adaptive card which is a form with product details and stock input, into the conversation.

The members in the team can then update stock information in the same conversation, by adding stock unit value in the input field and select **Submit**.

## Frameworks

![drop](https://img.shields.io/badge/Bot&nbsp;Framework-4.7-green.svg)


## Version history

Version|Date|Author|Comments
-------|----|----|--------
1.0|February 2022|Rabia Williams|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone the repository

    ```bash
    git clone https://github.com/OfficeDev/m365-msteams-northwind-app-samples
    ```

- In a console, navigate to `src/extend-with-capabilities/aad/MessagingExtension/`

    ```bash
    cd src/extend-with-capabilities/aad/MessagingExtension/
    ```

- Install modules

    ```bash
    npm install
    ```

- Run ngrok - point to port 3978

    ```bash
    ngrok http -host-header=rewrite 3978
    ```

- Since messaging extensions utilize the Azure Bot Framework, you will need to register a new bot. 
[These instructions](https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/create-a-bot-for-teams#register-your-web-service-with-the-bot-framework) provide options for registering with or without an Azure subscription. 
  - Be sure to enable the Microsoft Teams bot channel so your solution can communicate with Microsoft Teams
  - For local testing, set the messaging endpoint to the https URL returned by ngrok plus "/api/messages"
  - Note the bot's Application ID and password (also called the Client Secret) assigned to your bot during the registration process. In the Azure portal this is under the Bot Registration settings; in the legacy portal it's in the Settings tab. Click Manage to go to Azure AD to obtain the Client Secret. You may need to create a new Application Secret in order to have an opportunity to copy it out of the Azure portal. 

- Update the `.env` configuration for the bot to use the Microsoft App Id and App Password (aka Client Secret) from the previous step.
BOT_REG_AAD_APP_ID=&lt;Microsoft App Id&gt;
BOT_REG_AAD_APP_PASSWORD=&lt;Client Secret&gt;


  Upload the resulting zip file into Teams [using these instructions](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-upload).

- Run the bot locally
    ```bash
    npm start
    ```

- Test in Microsoft Teams by clicking the ... beneath the compose box in a Team where the application has been installed. Load the app.
- Search the product in the search box and select the desired product and post the card into the channel conversation.
- Update the stock value and select **Submit**

## Features

This is a simple Search messaging extension to show case how you can search and display data from a third party data base in Microsoft Teams. It also shows how you can update data in the third party data base all in the same conversation.

## Debug and test locally

TBD