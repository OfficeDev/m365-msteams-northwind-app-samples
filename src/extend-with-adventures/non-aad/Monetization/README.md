# 99 - Monetization

Changes:

1. Install projects from MonetizationCodeSample into Azure and register in AAD:
   - AppSourceMockWebApp
   - SaaSOfferMockData
   - SaaSSampleWebApi
   - SaaSSampleWebApp

2. Add GUID and scope (api://) pf tje SaaSSampleWebApi to .env file

3. Update the Teams app registration:
   - Expose the Teams app ID (from step 03-TeamSSOsApp) to the Web SaaSSampleWebApi app registration and authorize the scope
   - Add access_as_user permission to the SaaSSampleWebApi
   - Grant administrator consent to this permission

4. 