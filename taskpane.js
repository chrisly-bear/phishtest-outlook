/*
 * PhishTest Header Checker - Outlook Add-in
 *
 * Looks for the X-PHISHTEST header in the current email.
 * If found, displays its value as an ErrorMessage notification
 * (red banner) above the message body.
 *
 * NOTE: Office.js does not have a "WarningMessage" notification type.
 * The available types are: InformationalMessage, ErrorMessage,
 * ProgressIndicator, and InsightMessage. We use ErrorMessage as the
 * closest visual equivalent to a warning.
 */

const NOTIFICATION_KEY = "phishtest-warning";
const HEADER_NAME = "x-phishtest";
// Office.js notification messages have a 150-character limit.
const MAX_NOTIFICATION_LENGTH = 150;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        const checkButton = document.getElementById("checkButton");
        checkButton.disabled = false;

        // Wire the button click handler
        checkButton.addEventListener("click", checkHeaders);

        // Auto-check when the task pane loads
        checkHeaders();

        // Register for ItemChanged events so the pane re-checks
        // when the user selects a different message while pinned
        Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            () => {
                checkHeaders();
            }
        );
    }
});

/**
 * Main function: fetch headers, parse for X-PHISHTEST, show notification.
 */
function checkHeaders() {
    const statusEl = document.getElementById("status");
    const headerValueEl = document.getElementById("headerValue");
    const headerLabel = document.getElementById("headerLabel");

    // Reset UI
    setStatus(statusEl, "Checking headers...", "loading");
    headerValueEl.classList.remove("visible");
    headerLabel.style.display = "none";

    // Verify requirement set support at runtime
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
        setStatus(
            statusEl,
            "This add-in requires Outlook Mailbox API 1.8 or later. " +
            "Please update Outlook to the latest version.",
            "error"
        );
        return;
    }

    const item = Office.context.mailbox.item;

    if (!item || !item.getAllInternetHeadersAsync) {
        setStatus(
            statusEl,
            "Unable to access the current message. Make sure a message is selected.",
            "error"
        );
        return;
    }

    item.getAllInternetHeadersAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            setStatus(
                statusEl,
                "Failed to retrieve headers: " + (asyncResult.error ? asyncResult.error.message : "Unknown error"),
                "error"
            );
            return;
        }

        const rawHeaders = asyncResult.value || "";
        const phishTestValue = parseHeader(rawHeaders, HEADER_NAME);

        if (phishTestValue !== null) {
            // Found the X-PHISHTEST header - show warning notification
            // Handle empty value: an empty notification may not render,
            // so use a fallback message
            const displayValue = phishTestValue || "X-PHISHTEST header is present but empty.";
            showWarningNotification(displayValue);

            // Display full value in the task pane
            headerValueEl.textContent = phishTestValue;
            headerValueEl.classList.add("visible");
            headerLabel.style.display = "block";

            const truncatedNote =
                displayValue.length > MAX_NOTIFICATION_LENGTH
                    ? " (truncated in banner — full value shown below)"
                    : "";

            setStatus(
                statusEl,
                "X-PHISHTEST header found" + truncatedNote + ". Warning banner displayed above the message.",
                "warning"
            );
        } else {
            // No X-PHISHTEST header - remove any stale notification
            removeWarningNotification();

            setStatus(
                statusEl,
                "No X-PHISHTEST header found in this message.",
                "success"
            );
        }
    });
}

/**
 * Show (or replace) the warning notification on the current mail item.
 * Uses ErrorMessage type since WarningMessage does not exist in Office.js.
 * The notification text is truncated to 150 chars if needed.
 */
function showWarningNotification(headerValue) {
    const item = Office.context.mailbox.item;
    if (!item || !item.notificationMessages) {
        return;
    }

    // Truncate if the header value exceeds the 150-char notification limit
    let notificationText = headerValue;
    if (notificationText.length > MAX_NOTIFICATION_LENGTH) {
        notificationText = notificationText.substring(0, MAX_NOTIFICATION_LENGTH - 3) + "...";
    }

    const notificationDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: notificationText,
    };

    // Use replaceAsync to avoid errors if a notification with this key already exists
    item.notificationMessages.replaceAsync(
        NOTIFICATION_KEY,
        notificationDetails,
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error("Failed to show warning notification:", result.error);
            }
        }
    );
}

/**
 * Remove any existing PhishTest warning notification from the current item.
 */
function removeWarningNotification() {
    const item = Office.context.mailbox.item;
    if (!item || !item.notificationMessages) {
        return;
    }

    item.notificationMessages.removeAsync(NOTIFICATION_KEY, (result) => {
        // Ignore errors - the notification may not exist yet
        if (result.status === Office.AsyncResultStatus.Failed) {
            // Expected if no notification was previously added
            console.log("No existing notification to remove:", result.error ? result.error.message : "");
        }
    });
}

/**
 * Parse raw RFC 5322 headers and extract the value of a specific header.
 *
 * Handles:
 * - Case-insensitive header name matching
 * - RFC 5322 header folding (continuation lines starting with space/tab)
 * - Multiple headers with the same name (joins values with "; ")
 *
 * @param {string} rawHeaders - Full raw header string from getAllInternetHeadersAsync
 * @param {string} headerName - Header name to find (case-insensitive, without trailing colon)
 * @returns {string|null} The header value, or null if not found
 */
function parseHeader(rawHeaders, headerName) {
    if (!rawHeaders || !headerName) {
        return null;
    }

    // Unfold RFC 5322 continuation lines:
    // A line starting with space or tab is a continuation of the previous line.
    // Replace "newline + (space or tab)" with a single space to unfold.
    const unfolded = rawHeaders.replace(/\r?\n[ \t]+/g, " ");

    // Split into individual header lines
    const lines = unfolded.split(/\r?\n/);

    const target = headerName.toLowerCase();
    const values = [];

    for (const line of lines) {
        const colonIndex = line.indexOf(":");
        if (colonIndex === -1) {
            continue;
        }

        const name = line.substring(0, colonIndex).trim().toLowerCase();
        const value = line.substring(colonIndex + 1).trim();

        if (name === target) {
            values.push(value);
        }
    }

    if (values.length === 0) {
        return null;
    }

    // Join multiple values with "; " if the header appears more than once
    return values.join("; ");
}

// Helper to set status message
function setStatus(el, message, type) {
    el.textContent = message;
    el.className = "status " + type;
}
