---
page_type: sample
description: This sample demonstrates how to use the Microsoft Graph .NET SDK to access data in Office 365 from Microsoft Bot Framework bots.
products:
- ms-graph
- microsoft-graph-calendar-api
- office-exchange-online
languages:
- csharp
---
# Microsoft Graph Bot Framework sample

This sample demonstrates how to use the Microsoft Graph .NET SDK to access data in Office 365 from Microsoft Bot Framework bots.

## Prerequisites

- [.NET SDK](https://dotnet.microsoft.com/download) version 6

  ```bash
  # determine dotnet version
  dotnet --version
  ```

## To try this sample

- In a terminal, navigate to `GraphCalendarBot`

    ```bash
    # change into project folder
    cd GraphCalendarBot
    ```

- Run the bot from a terminal or from Visual Studio, choose option A or B.

  A) From a terminal

  ```bash
  # run the bot
  dotnet run
  ```

  B) Or from Visual Studio

  - Launch Visual Studio
  - File -> Open -> Project/Solution
  - Navigate to `GraphCalendarBot` folder
  - Select `GraphCalendarBot.csproj` file
  - Press `F5` to run the project

## Testing the bot using Bot Framework Emulator

[Bot Framework Emulator](https://github.com/microsoft/botframework-emulator) is a desktop application that allows bot developers to test and debug their bots on localhost or running remotely through a tunnel.

- Install the Bot Framework Emulator version 4.9.0 or greater from [here](https://github.com/Microsoft/BotFramework-Emulator/releases)

### Connect to the bot using Bot Framework Emulator

- Launch Bot Framework Emulator
- File -> Open Bot
- Enter a Bot URL of `http://localhost:3978/api/messages`

## Create a Bot Channels registration

Deploy the bot to Azure using the following procedures.

1. Open a browser and navigate to the [Azure Portal](https://portal.azure.com). Login using the account associated with your Azure subscription.

1. Select the upper-left menu, then select **Create a resource**.

1. On the **New** page, search for `Azure Bot` and select **Azure Bot**.

1. On the **Azure Bot** page, select **Create**.

1. Fill in the required fields, and leave **Messaging endpoint** blank. The **Bot handle** field must be unique. Be sure to review the different pricing tiers and select what makes sense for your scenario. If this is just a learning exercise, you may want to select the free option.

1. Change the “Type of App” to “Multi Tenant”.

1. For **Microsoft App ID**, select **Create new Microsoft App ID**.

1. Select **Review + create**. Once validation completes, select **Create**.

1. Once deployment has finished, select **Go to resource**.

1. Under **Settings**, select **Configuration**. Select the **Manage** link next to **Microsoft App ID**.

1. Select **New client secret**. Add a description and choose an expiration, then select **Add**.

1. Copy the client secret value before you leave this page. You will need it in the following steps.

    > [!IMPORTANT]
    > This client secret is never shown again, so make sure you copy it now. You will need to enter this value in multiple places so keep it safe.

1. Select **Overview** in the left-hand menu. Copy the value of the **Application (client) ID** and save it, you will need it in the following steps.

## Create a web app registration

1. Return to the home page of the Azure portal, then select **Azure Active Directory**.

1. Select **App registrations**.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `Graph Calendar Bot Auth`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Web` and set the value to `https://token.botframework.com/.auth/web/redirect`.

1. Select **Register**. On the **Graph Calendar Bot Auth** page, copy the value of the **Application (client) ID** and save it, you will need it in the following steps.

1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value in **Description** and select one of the options for **Expires** and select **Add**.

1. Copy the client secret value before you leave this page. You will need it in the following steps.

1. Select **API permissions**, then select **Add a permission**.

1. Select **Microsoft Graph**, then select **Delegated permissions**.

1. Select the following permissions, then select **Add permissions**.

    - **openid**
    - **profile**
    - **Calendars.ReadWrite**
    - **MailboxSettings.Read**

### About permissions

Consider what each of those permission scopes allows the bot to do, and what the bot will use them for.

- **openid** and **profile**: allows the bot to sign users in and get basic information from Azure AD in the identity token.
- **Calendars.ReadWrite**: allows the bot to read the user's calendar and to add new events to the user's calendar.
- **MailboxSettings.Read**: allows the bot to read the user's mailbox settings. The bot will use this to get the user's selected time zone.
- **User.Read**: allows the bot to get the user's profile from Microsoft Graph. The bot will use this to get the user's name.

## Add OAuth connection to the bot

1. Navigate to your bot's Azure Bot page in the Azure Portal. Select **Configuration** under **Settings**.

1. Select **Add OAuth Connection Settings**.

1. Fill in the form as follows, then select **Save**.

    - **Name**: `GraphBotAuth`
    - **Provider**: **Azure Active Directory v2**
    - **Client id**: The application ID of your **Graph Calendar Bot Auth** registration.
    - **Client secret**: The client secret of your **Graph Calendar Bot Auth** registration.
    - **Token Exchange URL**: Leave blank
    - **Tenant ID**: `common`
    - **Scopes**: `openid profile Calendars.ReadWrite MailboxSettings.Read User.Read`

1. Select the **GraphBotAuth** entry under **OAuth Connection Settings**.

1. Select **Test Connection**. This opens a new browser window or tab to start the OAuth flow.

1. If necessary, sign in. Review the list of requested permissions, then select **Accept**.

1. You should see a **Test Connection to 'GraphBotAuth' Succeeded** message.

> [!TIP]
> You can select the **Copy Token** button on this page, and paste the token into [https://jwt.ms](https://jwt.ms) to see the claims inside the token. This is useful when troubleshooting authentication errors.

## Test authentication

1. Save all of your changes and start the bot with `dotnet run`.

1. Open the Bot Framework Emulator. Select the gear icon &#9881; on the bottom left.

1. Enter the local path to your installation of ngrok, and enable the **Bypass ngrok for local addresses** and **Run ngrok when the Emulator starts up** options. Select **Save**.

1. Select the **File** menu, then **New Bot Configuration...**.

1. Fill in the fields as follows.

    - **Bot name:** `CalendarBot`
    - **Endpoint URL:** `https://localhost:3978/api/messages`
    - **Microsoft App ID:** the application ID of your **Graph Calendar Bot** app registration
    - **Microsoft App password:** your **Graph Calendar Bot** client secret
    - **Encrypt keys stored in your bot configuration:** Enabled

1. Select **Save and connect**. After the emulator connects, you should see `Welcome to Microsoft Graph CalendarBot. Type anything to get started.`

1. Type some text and send it to the bot. The bot responds with a login prompt.

1. Select the **Login** button. The emulator prompts you to confirm the URL that starts with `oauthlink://https://token.botframeworkcom`. Select **Confirm** to continue.

1. In the pop-up window, login with your Microsoft 365 account. Review the requested permissions and accept.

1. Once authentication and consent are complete, the pop-up window provides a validation code. Copy the code and close the window.

1. Enter the validation code in the chat window to complete the login.

1. If you select the **Show token** button (or type `show token`), the bot displays the access token. The **Log out** button (or typing `log out`) will log you out.

> [!TIP]
> You may receive the following error message in the Bot Framework Emulator when starting a conversation with the bot.
>
> ```text
> Failed to generate an actual sign-in link: Error: Failed to connect to ngrok instance for OAuth postback URL:
> FetchError: request to http://127.0.0.1:4041/api/tunnels failed, reason: connect ECONNREFUSED 127.0.0.1:4041
> ```
>
> If this happens, be sure to enable the **Run ngrok when the Emulator starts up** option in the emulator settings and restart the emulator.

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
