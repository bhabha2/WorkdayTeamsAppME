# Overview of Workday Teams Message Extension sample

The purpose of this sample app is to showcase how to build a message extension in a Teams App that can be used as a M365 Copilot Plugin.  
[Message extension](https://docs.microsoft.com/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions?tabs=nodejs) allows users to search for information in an external system and share results through a Copilot Prompt or the compose message area of the Microsoft Teams client. You can now build and run your search-based message extensions in Teams, Copilot for Windows desktop and web experiences.

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 16, 18
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
> - Access to Workday environment to create connected App
> - [Set up your dev environment for extending Teams apps across Microsoft 365](https://aka.ms/teamsfx-m365-apps-prerequisites)
>   Please note that after you enrolled your developer tenant in Office 365 Target Release, it may take couple days for the enrollment to take effect.
> - [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Teams Toolkit CLI](https://aka.ms/teams-toolkit-cli)
> - Microsoft 365 Copilot enabled environment.Login to Workday & search for Register API Client

Login to Workday & search for Register API Client
![alt text](image-1.png)
![alt text](image-2.png)
Register a new API client & Choose as guided below.

>Client Grant Type - Authorization Code Grant  
>Access Token Type - Bearer  
>Redirection URL: https://token.botframework.com/.auth/web/redirect  
>Non-Expiring Refresh Tokens - Configure as required  
>Grant Administrative Consent: Enabled  
>Scope(Functional Areas) - as required  

![alt text](image-3.png)

Once registered the Endpoints are displayed. Copy the Client ID & the endpoints as they will be needed to configure the OAuth connection in the Azure bot
![alt text](image-4.png)

Save & go back to the menu and search for Maintain API Client Access
![alt text](image-5.png)
Click on the popout icon next to the API Client created

![alt text](image-6.png)
The view API Client screen appears. Click on ... next to the API Client name at the top of the screen to display the menu with option to Generate New Client Secret.Please select the same. 
![alt text](image-7.png)
Select the Confirm option to generate new secret. Copy the secret as this will be needed to configure the OAuth connection in Azure.
![alt text](image.png)

2. **Visual Studio Code - Teams toolkit:**
   - Select the Teams Toolkit icon on the left in the VS Code toolbar.
   - In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.

3. **Login to Azure** & create a resource group to deploy the solution. Copy the subscription-id and resource group name.

4. **In Visual studio**, enter the values copied above in .env.dev file under env folder.
![alt text](image-51.png)

- Using the Teams Toolkit menu, under Lifecycle - choose each of the options Provision & Deploy sequentially.  
Provision will setup & fetch the necessary configuration parameters required to deploy the app in azure and  
Deploy will deploy the app & bot in Azure within the Resource group created in earlier step above.

- Wait for the deployment to be successfully completed.

5. **Login to Azure**  
- Click on the bot created under the resource group. Check for Azure bot in the resource type column.
![alt text](image-52.png)

- Select Configuration & click on `Add OAuth Connection Settings` button to enter the OAuth Connection details created in Step 1 & save the same.
![alt text](image-53.png)

- Go back to Overview view and select the Azure bot. Under Settings on the left menu, select Configuration & copy the values for `BOT_ID`, `BOT_PASSWORD` by clicking on the edit options
![alt text](image-54.png)

6. **Using Visual Studio** Create .env file under the src folder and include the values for MicrosoftAppId(`BOT_ID`), MicrosoftAppPassword (`BOT_PASSWORD`) & connectionName (`OAuth Connection Name`):  
![alt text](image-55.png)

- Redeploy the solution using the `Deploy` option from the Teams toolkit.

- Click on `Publish` option from the Teams toolkit to create the App Package file.  
   Publish will create the App Package file that includes the App manifest.json file. Confirming the publish will publish the latest app package to your Teams environment. Once published, the app will need to be approved by Teams Admin via the Teams Admin center.
   If you wish to publish the app in another environment, you can upload this package file there.

7. Using a Teams environment with access to upload custom apps, 
- upload the manifest.zip package found appPackage\build folder using the `Upload an app` option.
![alt text](image-56.png)

8. To trigger the app as Message Extension, `@mention` Your message extension from the `search box area`, `@mention` your message extension from the `compose message area` or click the `➕` under compose message area to find your message extension.

9. To trigger the Message Extension through Copilot as a Plugin, you can:  
   a. Open the `Copilot` app and enable the plugin from Plugin popup list.  
b. send a prompt to trigger your plugin.
> Note: This prompt may not always make Copilot include a response from your message extension. If it happens, try some other prompts or leave a feedback to us by thumbing down the Copilot response and leave a message tagged with [MessageExtension].  
      You may want to include a note that says current this only works in Copilot chat in Teams. It does not work in
https://copilot.microsoft.com
(Work or Web).

**Congratulations**! You can now search Salesforce opportunities in Teams via Copilot.

![Search ME Copilot](https://github.com/OfficeDev/TeamsFx/assets/107838226/a718b206-33ed-4d3e-99af-376c1f159c2b)

## What's included in the template

| Folder        | Contents                                     |
| ------------- | -------------------------------------------- |
| `.vscode/`    | VSCode files for debugging                   |
| `appPackage/` | Templates for the Teams application manifest |
| `env/`        | Environment files                            |
| `infra/`      | Templates for provisioning Azure resources   |
| `src/`        | The source code for the search application   |

The following files can be customized and demonstrate an example implementation to get you started.

| File               | Contents                                                                                       |
| ------------------ | ---------------------------------------------------------------------------------------------- |
| `src/searchApp.js` | Handles the business logic for this app template to query npm registry and return result list. |
| `src/index.js`     | `index.js` is used to setup and configure the Message Extension.                               |

The following are Teams Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Teams Toolkit works.

| File                 | Contents                                                                                                                                  |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `teamsapp.yml`       | This is the main Teams Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions. |
| `teamsapp.local.yml` | This overrides `teamsapp.yml` with actions that enable local execution and debugging.                                                     |
### Troubleshooting tip
If you're facing challenges while testing your plugin, you can enable 'developer mode'. Developer mode provides information on the plugin selected by the Copilot orchestrator to respond to the prompt. It also shows the available functions in the plugin and the API call's status code.
 
To enable developer mode, type the following into Copilot:
```
-developer on
```
For additional information on common problems and how to fix them, see the  [troubleshooting](Troubleshooting.md) guide.
 
Now just execute your prompt. This time, the output will look like this:
 
![The developer mode in action](./images/03-03b-developer-mode.png)
 
As you can notice, below the response generated by Copilot, we have a table that provides us insightful information about what happened behind the scenes:
 
- Under **Enabled plugins**, we can see that Copilot has identified that the Northwind Inventory plugin is enabled.
- Under **Matched functions**, we can see that Copilot has determined that the Northwind inventory plugin offers three functions: `inventorySearch`, `discountSearch`, and `companySearch`.
- Under **Selected functions for execution**, we can see that Copilot has selected the `inventorySearch` function to respond to the prompt.
- Under **Function execution details**, we can see some detailed information about the execution, like the HTTP response returned by the plugin to the Copilot engine.

## Extend the template

Following documentation will help you to extend the template.

- [Add or manage the environment](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Create multi-capability app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-capability)
- [Add single sign on to your app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-single-sign-on)
- [Access data in Microsoft Graph](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk#microsoft-graph-scenarios)
- [Use an existing Microsoft Entra application](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-existing-aad-app)
- [Customize the Teams app manifest](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- [Collaborate on app development](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Set up the CI/CD pipeline](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-cicd-template)
- [Publish the app to your organization or the Microsoft Teams app store](https://learn.microsoft.com/microsoftteams/platform/toolkit/publish)
- [Develop with Teams Toolkit CLI](https://aka.ms/teams-toolkit-cli/debug)
- [Preview the app on mobile clients](https://github.com/OfficeDev/TeamsFx/wiki/Run-and-debug-your-Teams-application-on-iOS-or-Android-client)
- [Extend Microsoft 365 Copilot](https://aka.ms/teamsfx-copilot-plugin)



