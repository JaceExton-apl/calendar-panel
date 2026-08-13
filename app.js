const CLIENT_ID =
    "b243d32a-4459-476c-8728-042bfb9bb8c5";

const TENANT_ID =
    "43001d36-133b-47ee-aaae-9fdd3a28fba8";


const msalConfig = {
    auth: {
        clientId: CLIENT_ID,

        authority:
            `https://login.microsoftonline.com/${TENANT_ID}`,

        redirectUri:
            window.location.origin +
            window.location.pathname
    },

    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false
    }
};


const loginRequest = {
    scopes: [
        "User.Read",
        "Calendars.Read"
    ],
    prompt: "select_account"
};


const msalInstance =
    new msal.PublicClientApplication(
        msalConfig
    );


const signInButton =
    document.getElementById(
        "signInButton"
    );

const signedOut =
    document.getElementById(
        "signedOut"
    );

const signedIn =
    document.getElementById(
        "signedIn"
    );


signInButton.addEventListener(
    "click",
    signIn
);


async function signIn() {

    try {

        const response =
            await msalInstance.loginPopup(
                loginRequest
            );

        msalInstance.setActiveAccount(
            response.account
        );

        signedOut.hidden = true;
        signedIn.hidden = false;

        await loadCalendar();

    } catch (error) {

        console.error(
            "Microsoft sign-in error:",
            error
        );

        alert(
            "Microsoft sign-in failed. Check the browser console."
        );
    }
}


async function getAccessToken() {

    let account =
        msalInstance.getActiveAccount();


    if (!account) {

        const accounts =
            msalInstance.getAllAccounts();

        if (accounts.length > 0) {

            account = accounts[0];

            msalInstance.setActiveAccount(
                account
            );
        }
    }


    if (!account) {
        return null;
    }


    try {

        const response =
            await msalInstance.acquireTokenSilent({
                scopes:
                    loginRequest.scopes,

                account:
                    account
            });

        return response.accessToken;

    } catch (error) {

        console.warn(
            "Silent token acquisition failed. Trying popup:",
            error
        );

        const response =
            await msalInstance.acquireTokenPopup({
                scopes:
                    loginRequest.scopes
            });

        return response.accessToken;
    }
}


async function loadCalendar() {

    const token =
        await getAccessToken();

    if (!token) {
        return;
    }


    const start =
        new Date();

    start.setHours(
        0,
        0,
        0,
        0
    );


    const end =
        new Date();

    end.setHours(
        23,
        59,
        59,
        999
    );


    const endpoint =
        "https://graph.microsoft.com/v1.0" +
        "/me/calendarView" +

        "?startDateTime=" +
        encodeURIComponent(
            start.toISOString()
        ) +

        "&endDateTime=" +
        encodeURIComponent(
            end.toISOString()
        ) +

        "&$orderby=start/dateTime";


    const response =
        await fetch(
            endpoint,
            {
                headers: {
                    Authorization:
                        `Bearer ${token}`,

                    Prefer:
                        `outlook.timezone="${
                            Intl.DateTimeFormat()
                                .resolvedOptions()
                                .timeZone
                        }"`
                }
            }
        );


    if (!response.ok) {

        const error =
            await response.text();

        console.error(
            "Microsoft Graph calendar error:",
            error
        );

        alert(
            "Unable to retrieve calendar."
        );

        return;
    }


    const data =
        await response.json();

    displayEvents(
        data.value
    );
}


function displayEvents(events) {

    const container =
        document.getElementById(
            "events"
        );

    container.innerHTML = "";


    const now =
        new Date();

    let nextEvent = null;


    for (const event of events) {

        const start =
            new Date(
                event.start.dateTime
            );

        const end =
            new Date(
                event.end.dateTime
            );


        if (
            !nextEvent &&
            end > now
        ) {
            nextEvent = event;
        }


        const eventRow =
            document.createElement(
                "div"
            );

        eventRow.className =
            "event";


        const time =
            document.createElement(
                "div"
            );

        time.className =
            "event-time";

        time.textContent =
            start.toLocaleTimeString(
                [],
                {
                    hour:
                        "2-digit",

                    minute:
                        "2-digit"
                }
            );


        const information =
            document.createElement(
                "div"
            );


        const title =
            document.createElement(
                "div"
            );

        title.className =
            "event-title";

        title.textContent =
            event.subject ||
            "Meeting";


        information.appendChild(
            title
        );


        if (
            event.location &&
            event.location.displayName
        ) {

            const location =
                document.createElement(
                    "div"
                );

            location.className =
                "event-location";

            location.textContent =
                event.location.displayName;

            information.appendChild(
                location
            );
        }


        eventRow.appendChild(
            time
        );

        eventRow.appendChild(
            information
        );

        container.appendChild(
            eventRow
        );
    }


    showNextMeeting(
        nextEvent
    );
}


function showNextMeeting(event) {

    const display =
        document.getElementById(
            "nextMeeting"
        );


    if (!event) {

        display.textContent =
            "No more meetings today";

        return;
    }


    const start =
        new Date(
            event.start.dateTime
        );


    display.textContent =
        `${start.toLocaleTimeString(
            [],
            {
                hour:
                    "2-digit",

                minute:
                    "2-digit"
            }
        )} — ${event.subject}`;
}


function updateClock() {

    const now =
        new Date();


    document.getElementById(
        "clock"
    ).textContent =
        now.toLocaleTimeString(
            [],
            {
                hour:
                    "2-digit",

                minute:
                    "2-digit"
            }
        );


    document.getElementById(
        "date"
    ).textContent =
        now.toLocaleDateString(
            [],
            {
                weekday:
                    "long",

                day:
                    "numeric",

                month:
                    "long"
            }
        );
}


updateClock();


setInterval(
    updateClock,
    1000
);


setInterval(
    loadCalendar,
    5 * 60 * 1000
);