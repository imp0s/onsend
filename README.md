# Word Attachment Cleaner on Send

This project contains an Outlook on-send add-in that inspects outgoing messages for Word attachments. When a `.docx` file is found, the add-in prompts the sender to remove metadata (e.g., author fields) and comments before delivery. If accepted, the attachments are cleaned and replaced; if declined, the message sends unmodified.

## Project layout

- `src/`: TypeScript sources for the add-in runtime and attachment cleanup helpers.
- `public/`: Dialog assets for user confirmation.
- `manifest/manifest.xml`: Mail add-in manifest configured for the item-send event.
- `.github/workflows/ci.yml`: GitHub Actions workflow to build and test.
- `tests/`: Unit tests for the document cleanup utilities.

## Prerequisites

- Node.js 18+
- npm
- Microsoft 365 subscription with the ability to sideload Outlook add-ins.

> Note: The Office.js runtime is loaded from Microsoft's CDN inside the dialog page rather than installed from npm.

## Clone and configure for Outlook

1. Clone the repository (replace the URL with your fork if needed):

   ```bash
   git clone https://github.com/imp0s/onsend.git
   cd onsend
   sudo npm install
   ```

2. Build the add-in once to produce the compiled runtime in `dist/`:

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

   # Serve the repo over HTTPS on port 3000
   npx http-server -S -C certs/localhost.crt -K certs/localhost.key -p 3000 .
   ```

4. The manifest already includes default URLs that point to `https://localhost:3000` for the task pane (`public/commands.html`), function file (`public/functions.html`), dialog, and runtime script. If you host on a different origin/port, update [`manifest/manifest.xml`](manifest/manifest.xml) to match.

5. Sideload the XML add-in manifest (not the unified Microsoft 365 manifest) into Outlook (desktop or web) using Microsoft’s [add-in only manifest sideloading guidance](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

6. Compose a message with one or more `.docx` attachments. When you send the message, the add-in prompts to clean the attachments.

> Example files referenced above:
> - Manifest definition: [`manifest/manifest.xml`](manifest/manifest.xml)
> - On-send handler wiring: [`src/addin.ts`](src/addin.ts)
> - User confirmation dialog: [`public/dialog.html`](public/dialog.html)

## Running locally

1. Build the project: `npm run build`.
2. Serve the repository root over HTTPS on port 3000 (required by Office add-ins). Any static server works; for example:

   ```bash
   # Generate a self-signed certificate valid for localhost (adjust -subj if needed)
   mkdir -p certs
   openssl req -x509 -nodes -days 365 -newkey rsa:2048 \
     -keyout certs/localhost.key -out certs/localhost.crt \
     -subj "/CN=localhost"

   # Serve the repo over HTTPS on port 3000
   npx http-server -S -C certs/localhost.crt -K certs/localhost.key -p 3000 .
   ```

3. Update `manifest/manifest.xml` URLs if you host on a different origin/port.
4. Sideload the manifest into Outlook (desktop or web) following Microsoft's [sideload guidance](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).
5. Compose a message with one or more `.docx` attachments. When you send the message, the add-in prompts to clean the Word attachments. Choosing **Yes** removes metadata and comments and resaves the attachments before the send completes.

## Testing

Make sure dependencies (including Jest) are installed, then run the unit tests with coverage:

```bash
sudo npm install
npm test
```

> If `npm test` reports that `jest` is missing, reinstall dependencies with `sudo npm install` to ensure dev dependencies are available.

## How it works

- The on-send handler (`onMessageSend` in `src/addin.ts`) enumerates attachments and filters for `.docx` files.
- The user is prompted via an Office dialog (backed by `public/dialog.html`) to decide whether to clean the files.
- If confirmed, each Word attachment is downloaded, cleaned with `removeMetadataAndComments` from `src/docCleanup.ts`, reattached, and the send continues.
- The cleanup routine strips common core properties (creator, modified by, created/modified timestamps, last printed) and removes all Word comments by clearing `word/comments.xml`.

## GitHub Actions

The workflow at `.github/workflows/ci.yml` installs dependencies, builds the TypeScript sources, and executes the Jest test suite to keep the project healthy.

## Getting more testing feedback sooner

To avoid waiting for the GitHub Action to finish, ask Codex (or your local environment) to run the same checks before pushing:

- Run the unit suite locally: `npm test`.
- Rebuild after changes to catch TypeScript errors early: `npm run build`.
- If you use Codex as a coding assistant, explicitly request it to execute these commands and report results before finalizing its response. That way you see failures immediately instead of waiting for CI.
