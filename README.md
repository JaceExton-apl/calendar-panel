# Microsoft 365 Calendar Panel

A lightweight, fullscreen Microsoft 365 calendar display designed for Android tablets.

This project turns an inexpensive or repurposed Android tablet into a dedicated Microsoft 365 calendar panel, similar in concept to a meeting-room scheduling display.

The application is a static Progressive Web App (PWA) hosted using GitHub Pages. It authenticates users with Microsoft Entra ID using MSAL.js and retrieves calendar information directly from Microsoft Graph.

## Features

- Microsoft 365 sign-in
- Microsoft Graph calendar integration
- Displays today's calendar events
- Displays the next upcoming meeting
- Shows meeting times and locations
- Live clock and date
- Automatically refreshes calendar information
- Retains the signed-in Microsoft account between page refreshes
- Installable as a Progressive Web App
- Fullscreen display on supported Android devices
- No browser address bar when launched as an installed PWA
- Responsive interface suitable for small Android tablets
- Read-only calendar access

## Architecture

The application works as follows:

```text
Android Tablet
      │
      ▼
Calendar Panel PWA
      │
      ▼
GitHub Pages
      │
      ├──── Microsoft Entra ID
      │        Authentication
      │
      ▼
Microsoft Graph API
      │
      ▼
Microsoft 365 Calendar
```

GitHub Pages only hosts the static HTML, CSS and JavaScript files.

Calendar information is retrieved directly by the authenticated browser from Microsoft Graph.

## Requirements

You will need:

- A Microsoft 365 account
- Access to Microsoft Entra App Registrations
- A GitHub account
- A modern web browser
- An Android tablet for the finished display

The project was designed and tested around Android 12, although other modern Android versions should also work.

## Repository Structure

```text
calendar-panel/
│
├── index.html
├── app.js
├── style.css
├── manifest.json
├── icon-192.png
├── icon-512.png
└── README.md
```

## Microsoft Entra Setup

### 1. Create an App Registration

Open the Microsoft Entra admin centre.

Navigate to:

```text
Identity
→ Applications
→ App registrations
→ New registration
```

Create an application such as:

```text
Microsoft 365 Calendar Panel
```

For an application intended only for one Microsoft 365 organisation, select:

```text
Accounts in this organizational directory only
```

After registration, make a note of:

```text
Application (client) ID
Directory (tenant) ID
```

These values are required by `app.js`.

They are identifiers and are not passwords or client secrets.

## Authentication Platform

Under:

```text
App registrations
→ Calendar Panel
→ Authentication
```

add the platform:

```text
Single-page application
```

For local development, add:

```text
http://localhost:5500/
```

When the application is published using GitHub Pages, also add the production address.

For example:

```text
https://YOUR-GITHUB-USERNAME.github.io/calendar-panel/
```

The URL must match the actual GitHub Pages URL.

## Microsoft Graph Permissions

Navigate to:

```text
API permissions
→ Add a permission
→ Microsoft Graph
→ Delegated permissions
```

Add:

```text
Calendars.Read
```

This allows the application to read the signed-in user's calendar.

The application does not require `Calendars.ReadWrite` unless calendar editing functionality is added in the future.

Depending on the Microsoft 365 organisation's policies, administrator consent may be required.

## Configure app.js

Enter the Microsoft Entra Application ID and Tenant ID near the beginning of `app.js`.

For example:

```javascript
const CLIENT_ID =
    "YOUR-APPLICATION-CLIENT-ID";

const TENANT_ID =
    "YOUR-DIRECTORY-TENANT-ID";
```

Do not add a Microsoft Entra client secret to this project.

Browser-based Single Page Applications are public clients and must not contain client secrets.

## Authentication

Authentication is handled using Microsoft Authentication Library (MSAL).

The application requests the delegated Microsoft Graph permission:

```text
Calendars.Read
```

MSAL stores its authentication cache in browser `localStorage`.

This allows the application to restore the previously authenticated Microsoft account when the page reloads.

The application attempts silent token acquisition before requiring another interactive Microsoft sign-in.

## Calendar Data

Calendar information is retrieved from Microsoft Graph using:

```text
/me/calendarView
```

The application requests events occurring during the current day.

This allows recurring calendar events and individual meeting occurrences to be displayed correctly.

Calendar data is retrieved directly by the browser after Microsoft authentication.

It is not stored in the GitHub repository.

## Automatic Refresh

The application automatically refreshes Microsoft Graph calendar information every five minutes.

For example:

```javascript
setInterval(
    loadCalendar,
    5 * 60 * 1000
);
```

The page itself does not need to be manually refreshed to receive normal calendar updates.

## Running Locally

The application should be served through a local web server rather than opening `index.html` directly from Finder or File Explorer.

For example, using Visual Studio Code and Live Server:

```text
http://localhost:5500/
```

Make sure this address has also been registered as a Single-page application redirect URI in Microsoft Entra.

## Publishing with GitHub Pages

Open the repository settings:

```text
Settings
→ Pages
```

Under **Build and deployment**, select:

```text
Source:
Deploy from a branch
```

Then select:

```text
Branch:
main

Folder:
/ (root)
```

GitHub will publish the application at an address similar to:

```text
https://YOUR-GITHUB-USERNAME.github.io/calendar-panel/
```

Remember to add the exact published URL to the Microsoft Entra application's Single-page application redirect URIs.

## Installing on Android

Open the published GitHub Pages URL using Chrome on the Android tablet.

Confirm that Microsoft authentication and calendar retrieval are working.

Then use:

```text
Chrome
→ Menu
→ Install app
```

or:

```text
Chrome
→ Menu
→ Add to Home screen
```

depending on the version of Chrome.

Launch the resulting **Calendar Panel** icon from the Android home screen.

Do not launch the application from a normal Chrome browser tab if fullscreen PWA behaviour is required.

## Fullscreen Mode

The application uses a PWA manifest containing:

```json
"display": "fullscreen"
```

When installed as a PWA and launched from its Android application icon, supported browsers can display the calendar without the normal browser address bar and browser controls.

This provides a display similar to a dedicated meeting-room calendar panel.

Android system navigation controls may still be available unless additional Android kiosk configuration is used.

## Keeping the Tablet Awake

On many Android devices, Developer Options can be enabled by navigating to:

```text
Settings
→ About tablet
→ Build number
```

Tap **Build number** seven times.

Then locate:

```text
Developer options
→ Stay awake
```

Enable the option that keeps the screen awake while the tablet is connected to power.

The exact location of this setting varies between Android manufacturers.

## Security

This project intentionally does not contain:

- Microsoft passwords
- Microsoft access tokens in source code
- Microsoft refresh tokens in source code
- Microsoft Entra client secrets
- private keys
- certificates
- GitHub access tokens

The following values are present in the browser application:

```text
Application (client) ID
Directory (tenant) ID
```

These are application/tenant identifiers and are not authentication secrets.

### Authentication Cache

MSAL uses browser `localStorage` to retain authentication state.

This is useful for a dedicated calendar panel because it prevents users from having to authenticate after every page refresh.

However, physical access to an authenticated tablet should therefore be considered when deploying the application.

For public or shared installations, additional Android kiosk/device-management controls should be considered.

### Calendar Privacy

The application may display meeting subjects and locations.

Care should therefore be taken when using the panel in public areas.

A future version may optionally hide the subject and location of private calendar events.

## Troubleshooting

### Sign in button does nothing

Open the browser developer console and check for JavaScript errors.

If you see:

```text
msal is not defined
```

verify that the MSAL library is loaded in `index.html` before `app.js`.

### Redirect URI error

If Microsoft reports a redirect URI mismatch, verify the exact URL under:

```text
Microsoft Entra
→ App registrations
→ Calendar Panel
→ Authentication
→ Single-page application
```

For local development this may be:

```text
http://localhost:5500/
```

For GitHub Pages it may be:

```text
https://YOUR-GITHUB-USERNAME.github.io/calendar-panel/
```

The registered address must match the address used by the application.

### Wrong Microsoft account

The application can request Microsoft's account selector during authentication.

The login request can include:

```javascript
prompt: "select_account"
```

This allows another Microsoft account to be selected.

### Login is lost after refreshing

Verify that the MSAL configuration contains:

```javascript
cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
}
```

The application should also check `msalInstance.getAllAccounts()` when starting and attempt to restore the existing authenticated account.

### GitHub Pages appears to be running old code

GitHub Pages and browsers may temporarily cache static files.

Confirm that the latest GitHub Pages deployment has completed.

If necessary, version a JavaScript reference in `index.html`, for example:

```html
<script src="app.js?v=2"></script>
```

### Address bar still appears on Android

Opening the GitHub Pages URL directly in Chrome will still display Chrome's browser interface.

Install the application using:

```text
Add to Home screen
```

or:

```text
Install app
```

Then launch **Calendar Panel** using its Android home-screen icon.

## Current Limitations

The current version is intended primarily as a read-only personal Microsoft 365 calendar display.

Potential future improvements include:

- Available / Busy status
- Current meeting display
- Countdown to next meeting
- Microsoft Teams Join button
- Hide completed meetings
- Private meeting protection
- Tomorrow view
- Meeting-room resource calendars
- Improved offline/error handling
- Touch-optimised navigation
- Android kiosk lockdown
- Automatic launch after device restart
- Configurable themes and layouts

## Disclaimer

This is an independent project using Microsoft Graph and Microsoft Entra ID.

It is not an official Microsoft product and is not affiliated with or endorsed by Microsoft.

Microsoft, Microsoft 365, Microsoft Graph, Microsoft Entra and Microsoft Teams are trademarks of Microsoft Corporation.

## License

No licence has been specified for this project.

If the repository is intended to be shared or reused publicly, consider adding an appropriate open-source licence such as MIT.
