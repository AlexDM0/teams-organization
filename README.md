# Microsoft Teams Organization Tree Extractor

A fast, Typescript/Bun-based script that extracts the organization tree from Microsoft Teams (Entra ID) using the Microsoft Graph API and generates a beautiful, interactive D3.js visualization.

## Features
- **Device Code Flow Authentication**: Securely authenticates users without storing passwords.
- **Deep Graph API Integration**: Automatically resolves managers and fetches high-quality profile pictures encoded directly to Base64.
- **Interactive Visualization**: Uses `d3-org-chart` to render an aesthetic, self-contained interactive web UI for exploring the organizational hierarchy.

## Prerequisites
- **Bun**: Install [Bun](https://bun.sh/) if you haven't already.
- **Azure AD App Registration**: You need an application registered in Azure AD to authenticate the script. See the setup guide below.

## Azure AD Setup Guide

To get the `CLIENT_ID` and `TENANT_ID` for your `.env` file, follow these steps:

### 1. Register the Application
1. Go to the [Microsoft Entra admin center](https://entra.microsoft.com/) (or the [Azure Portal](https://portal.azure.com/)) and log in with your company account.
2. In the left-hand menu, expand **Applications** and click on **App registrations**.
3. Click **New registration** at the top.
4. Name it something recognizable, like *"Teams Org Tree Extractor"*.
5. Under **Supported account types**, leave the default ("Accounts in this organizational directory only").
6. Click **Register** at the bottom.

### 2. Copy your IDs
Immediately after registering, you will be taken to the app's **Overview** page. Look at the "Essentials" section at the top:
- Copy the **Application (client) ID** and paste it as your `CLIENT_ID` in the `.env` file.
- Copy the **Directory (tenant) ID** and paste it as your `TENANT_ID` in the `.env` file.

### 3. Enable the Device Code Flow (Crucial!)
Since we are using the interactive CLI login flow, we must explicitly allow it.
1. On the left menu of your App Registration, click **Authentication**.
2. Scroll all the way down to **Advanced settings**.
3. Look for **Allow public client flows** (or "Enable the following mobile and desktop flows") and toggle it to **Yes**.
4. Click **Save** at the top.

### 4. Grant API Permissions
Lastly, we need to give the app permission to read the organization's user profiles.
1. On the left menu, click **API permissions**.
2. Click **Add a permission** -> **Microsoft Graph** -> **Delegated permissions**.
3. In the search box, type `User.Read.All` and check the box next to it.
4. Click **Add permissions** at the bottom.
5. **Important:** Back on the API permissions screen, click the button that says **Grant admin consent for [Your Organization]** (you may need a company administrator to click this for you if you don't have admin rights).

## Quick Start

1. **Install Dependencies** (if not already installed)
   ```bash
   bun install
   ```

2. **Configure Environment**
   Copy `.env.example` to `.env` and fill in your `CLIENT_ID` and `TENANT_ID`:
   ```bash
   cp .env.example .env
   ```

3. **Run the Extractor**
   Use the start command to authenticate and fetch the data. This will generate `org_tree.json`.
   ```bash
   bun run start
   ```

4. **View the Visualization**
   Due to browser security policies (CORS), you must serve the `index.html` file using a local web server to view the generated JSON. You can do this easily with the serve command:
   ```bash
   bun run serve
   ```
   Open the printed URL in your browser to explore the interactive chart!
