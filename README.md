# Outlook Email Safety Checker

An Outlook on-send add-in that blocks emails with attachments and prevents sending to recipients outside your domain. Can be used if your organisation is worried about not having support for tools such as Tessian, Metadact within the Microsoft 365 service.

A simple tool to protect against sending emails outside the organisation or with attachments containing inappropriate metadata or comments.

## Project layout

- `src/`: JavaScript sources for the add-in runtime and configuration.
- `public/`: Assets loaded by the add-in runtime.
- `manifest/manifest.xml`: Mail add-in manifest configured for the item-send event.
- `.github/workflows/ci.yml`: GitHub Actions workflow to build and test.
- `tests/`: Unit tests for the safety helpers.

## Prerequisites

- Node.js 18+
- npm
- Microsoft 365 subscription with the ability to sideload Outlook add-ins.

> Note: The Office.js runtime is loaded from Microsoft's CDN inside the functions page rather than installed from npm.

## Clone and configure for Outlook

1. Clone the repository (replace the URL with your fork if needed):

   ```bash
   git clone https://github.com/imp0s/onsend.git
   cd onsend
   npm install
   ```

2. Build the add-in once to produce the compiled runtime in `public/app.js`:

   ```bash
   npm run build
   ```

3. Host the project root over HTTPS on `https://localhost:3000` (Office add-ins require HTTPS even during local sideloading). A simple option is [`http-server`](https://www.npmjs.com/package/http-server). If you do not already have a development certificate, create one and start the server:

   ```bash
   # Generate a self-signed certificate valid for localhost (adjust -subj if needed)
   mkdir -p certs
   openssl req -x509 -nodes -days 365 -newkey rsa:2048 \
     -keyout certs/localhost.key -out certs/localhost.crt \
     -subj "/CN=localhost"

   # Serve the repo over HTTPS on port 3000 with no caching
   npx http-server -S -C certs/localhost.crt -K certs/localhost.key -p 3000 . \
     -c-1 \
     -H "Cache-Control: no-store, must-revalidate" \
     -H "Pragma: no-cache" \
     -H "Expires: 0"
   ```

4. The manifest includes default URLs that point to `https://localhost:3000` for the function file (`public/functions.html`) and runtime script. If you host on a different origin/port, update [`manifest/manifest.xml`](manifest/manifest.xml) to match.

5. Sideload the XML add-in manifest into Outlook (desktop or web) using Microsoft’s [add-in only manifest sideloading guidance](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing). This effectively means:
   1. Visit https://outlook.office365.com/mail/inclientstore
   1. Log in
   1. Add your custom add-in by uploading the `metadata.xml` file.

6. Compose a message. The add-in blocks the send if any attachments are present or if recipients use domains outside the allowed list (the sender domain by default).

## Running locally

1. Build the project: `npm run build`.
2. Serve the repository root over HTTPS on port 3000 (required by Office add-ins). Any static server works; for example:

   ```bash
   # Generate a self-signed certificate valid for localhost (adjust -subj if needed)
   mkdir -p certs
   openssl req -x509 -nodes -days 365 -newkey rsa:2048 \
     -keyout certs/localhost.key -out certs/localhost.crt \
     -subj "/CN=localhost"

   # Serve the repo over HTTPS on port 3000 with no caching
   npx http-server -S -C certs/localhost.crt -K certs/localhost.key -p 3000 . \
     -c-1 \
     -H "Cache-Control: no-store, must-revalidate" \
     -H "Pragma: no-cache" \
     -H "Expires: 0"
   ```

3. Update `manifest/manifest.xml` URLs if you host on a different origin/port.
4. Sideload the manifest into Outlook (desktop or web) following Microsoft's [sideload guidance](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing). This effectively means:
   1. Visit https://outlook.office365.com/mail/inclientstore
   1. Log in
   1. Add your custom add-in by uploading the `metadata.xml` file.
5. Compose a message. The add-in stops sends that include attachments or recipients outside the permitted domain.

## Configuring allowed domains

By default, the add-in uses the sender's domain (from the signed-in account) to enforce recipient matching. Domains defined in [`src/config.js`](src/config.js) are appended to the 'allow list'. Update the `allowedDomainExtensions` array with your approved domains or suffixes.

## Testing

Make sure dependencies (including Jest) are installed, then run the unit tests with coverage:

```bash
npm install
npm test
```

## How it works

- The on-send handler (`onMessageSend` in `src/addin.js`) collects attachments and recipient addresses before the send.
- If the sender domain is known, all recipients must match it. Otherwise the domains listed in `allowedDomainExtensions` are used as suffix checks.
- Any attachment presence or domain mismatch shows an informational notification and blocks the send.

## GitHub Actions

The workflow at `.github/workflows/ci.yml` installs dependencies, builds the JavaScript bundle, and executes the Jest test suite to keep the project healthy.
