# Phishing Campaign Detector — Outlook Add-in

An Outlook add-in that inspects the headers of the currently selected email and looks for an `X-PHISHTEST` header. If found, it displays the header's contents as a red warning banner (notification) above the message body in the Outlook reading pane.

## Important note on notification types

The Office.js API **does not provide a `WarningMessage` notification type**. The available notification types are:

| Type | Visual | API set |
|---|---|---|
| `InformationalMessage` | Blue info banner with icon | 1.3 |
| `ErrorMessage` | Red error banner with Dismiss | 1.3 |
| `ProgressIndicator` | Progress spinner | 1.3 |
| `InsightMessage` | Info banner with action button | 1.10 |

This add-in uses **`ErrorMessage`** (red banner) as the closest visual equivalent to a warning. The banner includes a Dismiss button and can be set to persist across message changes.

Sources: [Microsoft Learn — ItemNotificationMessageType enum](https://learn.microsoft.com/en-us/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype), [Microsoft Learn — Create notifications](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/notifications)

## How it works

1. The add-in activates in **Message Read** mode (when a message is opened or selected in the reading pane).
2. When the task pane opens, it automatically calls `Office.context.mailbox.item.getAllInternetHeadersAsync()` to fetch the full RFC 5322 headers.
3. It parses the headers (case-insensitive, with RFC 5322 folding support) looking for `X-PHISHTEST`.
4. If found:
   - The header value is displayed as an `ErrorMessage` notification (red banner) above the message body using `notificationMessages.replaceAsync()`.
   - The full header value is also shown in the task pane.
   - If the value exceeds 150 characters (the notification text limit), it is truncated in the banner but shown in full in the task pane.
5. If not found, any previous PhishTest notification is removed from the item.
6. The task pane supports **pinning** — when pinned, it automatically re-checks when you select a different message (via the `ItemChanged` event).

## Requirements

- **Outlook for Mac** 16.38.506+ (new Mac UI) — requires Mailbox API 1.8
- **Outlook on Windows** — version 1910 (build 12130.20272) or later
- **Outlook on the web** — supported
- **Microsoft 365 subscription** (for Mac; classic COM/VSTO plugins are not supported)
- **Node.js** 18+ (for running the local dev server and tests)

Source: [Microsoft Learn — Outlook API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)

## Project structure

```
phishtest-addin/
├── manifest.xml          # Add-in manifest (classic XML format)
├── taskpane.html         # Task pane UI
├── taskpane.js           # Header parsing and notification logic
├── package.json          # Scripts and dependencies
├── assets/
│   ├── icon-16.png       # 16x16 icon
│   ├── icon-32.png       # 32x32 icon
│   └── icon-80.png       # 80x80 icon
├── test/
│   └── parser.test.js    # Unit tests for header parser
└── README.md             # This file
```

## Setup and local testing

### 1. Install dependencies

```bash
cd phishtest-addin
pnpm install
```

### 2. Run the parser unit tests

```bash
pnpm test
```

This runs 14 tests covering: simple header matching, case-insensitivity, RFC 5322 header folding, multiple headers with the same name, empty values, whitespace handling, and more.

### 3. Start the local HTTPS server

Outlook add-ins require HTTPS. You need a self-signed certificate:

```bash
# Generate a self-signed certificate
# NOTE: the Subject Alternative Name is required — Outlook's webview rejects
# certificates that only have CN=localhost and no SAN.
mkdir -p certs
openssl req -x509 -newkey rsa:2048 -keyout certs/key.pem \
  -out certs/cert.pem -days 365 -nodes \
  -subj "/CN=localhost" \
  -addext "subjectAltName=DNS:localhost,IP:127.0.0.1"

# Trust the certificate in the macOS login keychain
security add-trusted-cert -r trustRoot -k ~/Library/Keychains/login.keychain-db certs/cert.pem

# Start the HTTPS server
pnpm start
```

The task pane will be served at `https://localhost:3000/taskpane.html`.

### 4. Sideload the add-in on Outlook for Mac

1. Open **Outlook for Mac**.
2. Click **...** (more options) in the toolbar → **Get Add-ins**.
3. Click **Add from File...** (or **My Add-ins** → **Add from File...**).
4. Select the `manifest.xml` file from this project.
5. The add-in should appear in your add-ins list.

> **Note:** "Add from URL" is no longer supported on Mac. You must sideload via the manifest file.

Source: [Microsoft Learn — Sideload Outlook add-ins on Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-in-for-testing)

### 5. Test the add-in

1. Open or select an email in Outlook.
2. Click the **Check Headers** button in the ribbon to open the task pane.
3. The task pane automatically checks the current message on open. You can also click **Check Current Message** to re-check manually.
4. If the email has an `X-PHISHTEST` header, a red warning banner appears above the message body with the header's value.
5. The task pane also shows the full header value.

### 6. Pin the task pane (optional but recommended)

Click the pin icon on the task pane to keep it open. When pinned, it automatically re-checks each message as you select different emails in the list.

### 7. Clear the Outlook add-in cache

To make sure Outlook uses the latest version of the add-in, you can clear its add-in cache by removing the following files. Make sure that Outlook is not running.

```bash
rm -rf ~/Library/Containers/com.Microsoft.OsfWebHost/Data
rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef
rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/WebKit
rm -rf ~/Library/WebKit/com.microsoft.Outlook
```

## How to add an X-PHISHTEST header for testing

To test with a real email that has the `X-PHISHTEST` header, you can send yourself a test email with a tool like `swaks` or `curl`:

```bash
# Using swaks (Swiss Army Knife for SMTP)
swaks --to you@example.com --from test@example.com \
  --server your-smtp-server \
  --header "X-PHISHTEST: This is a simulated phishing warning"

# Or using curl with an SMTP relay
curl --url 'smtp://your-smtp-server:587' \
  --mail-from 'test@example.com' \
  --mail-rcpt 'you@example.com' \
  --upload-file - << 'EOF'
From: test@example.com
To: you@example.com
Subject: PhishTest
X-PHISHTEST: This is a simulated phishing warning

This is a test email body.
EOF
```

## API reference

- [`getAllInternetHeadersAsync`](https://learn.microsoft.com/en-us/javascript/api/outlook/office.messageread?view=outlook-js-preview#outlook-office-messageread-getallinternetheadersasync-member) — Mailbox 1.8, Read mode only
- [`notificationMessages.replaceAsync`](https://learn.microsoft.com/en-us/javascript/api/outlook/office.notificationmessages?view=outlook-js-preview) — Mailbox 1.3, Read and Compose
- [`ItemNotificationMessageType`](https://learn.microsoft.com/en-us/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) — enum for notification types
- [Outlook API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) — client support matrix

## License

MIT
