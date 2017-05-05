# Sample JavaScript SPA app that uses the Outlook REST APIs

This sample app is the result of following the [JavaScript tutorial on the Outlook Dev Center](https://docs.microsoft.com/en-us/outlook/rest/javascript-tutorial). If you follow that tutorial, you should come up with something fairly close to the code in this repository.

> **NOTE:** Looking for the version of this tutorial that used the Outlook API directly instead of Microsoft Graph? Check out the `outlook-api` branch. Note that Microsoft recommends using the Microsoft Graph to access mail, calendar, and contacts. You should use the Outlook APIs directly (via https://outlook.office.com/api) only if you require a feature that is not available on the Graph endpoints.

## Running the sample

If you didn't follow the tutorial and just want to download this repo and try it out, you need to do a few things first.

### Set up a web server

You can host these files on any web server you like. If you don't have a web server handy, the tutorial has steps to install [http-server](https://www.npmjs.com/package/http-server) to quickly create a development web server on the command line.

### Register the app

> **Note**: These steps tell you to use `http://localhost:8080` as your redirect URI. This is assuming that you are using `http-server` as specified in the tutorial. If you're using a different web server, be sure to change the redirect URI to the acutal URI used to access your app.

Head over to the Application Registration Portal to quickly get a client ID and secret. If you're not asked to sign in, click the **Go to app list** button and sign in with either your Microsoft account (Outlook.com), or your work or school account (Office 365). Once you're signed in, click the **Add an app** button. Enter 'javascript-tutorial` for the name and click **Create application**. 

Locate the **Platforms** section, and click **Add Platform**. Choose **Web**, then enter `http://localhost:8080` under **Redirect URIs**. Click **Save** to complete the registration. Copy the **Application Id** and save it. We'll need those values soon.

Here's what the details of your app registration should look like when you are done.

![The completed app registration](./readme-images/app-registration.PNG)

Replace the `YOUR APP ID HERE` value in `outlook-demo.js` with the application ID you generated as part of the app registration process.