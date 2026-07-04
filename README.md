# Empower Poster

Browser extension that uploads holdings and asset classification data
from [Empower](https://ira.empower-retirement.com/) (formerly Personal
Capital) to a user-specified external HTTPS endpoint. For example, you
could update a Google Sheets spreadsheet with the data by having the endpoint be
an Apps Script web app.

## Endpoint request/response API

The complete definitions of the types referenced below are in the
[`types` package](./packages/types/index.d.ts). It is available as a
Node package,
[@righteffort/empower-poster-types](https://www.npmjs.com/package/@righteffort/empower-poster-types).

### Request

The extension makes HTTPS POST requests with `Content-Type:
application/json`. The payload conforms to `PostPayload` and contains:
- `version`: API version.
- `holdings`: Array of `HoldingEntry` objects with `cusip`,
  `userAccountId`, `ticker`, `price`, `quantity`, `value`, and
  `fundFees`. `cusip` will generally be blank for manual holdings, and
  non-blank for others.
- `classifications`: `Classifications` object mapping tickers to asset
  classifications.
  - For each ticker there are one or more Classification entries, with
  fractions that add to 1.
  - Each Classification has the asset class, asset subclass (possibly
  empty), and fraction.
- accounts: Array of `Account` objects with `id`, `name`, and other
  details about the account. `id` is referenced by
  `HoldingEntry.userAccountId` and is unique in the array.

### Response

The response should have `Content-Type: application/json`, and the
body should conform to one of the response types:
- `SuccessResponse`: `{success: true, message?: string}`.
- `ErrorResponse`: `{success: false, error: string}`.

**Response Handling:**

The extension interprets the response to the POST request as follows:
- Non-2xx HTTPS status codes are treated as failures.
- Content-Type `application/json`: success based on the boolean `success` field in the JSON response body.
- Other values of Content-Types: empty body indicates success, non-empty is a failure message.

## Using in Google Sheets Apps Script

> [!IMPORTANT]
> Whenever you make changes to your Apps Script code, you must deploy
> the new version of the code -- see [Deployment
> Management](#deployment-management) below.

### Basic example

You can start by making a copy of [this
spreadsheet](https://docs.google.com/spreadsheets/d/1UlLLp5KLMbtYUJEV35B92Rj7j8ivbopPp-INh4W8lYw/edit),
which simply copies each of `holdings`, `classifications`, and
`accounts` to a separate sheet. The code is compiled from the source
in the [apps-script-minimal](./packages/apps-script-minimal) package.

### Advanced example: Updating Portfolio Tracker by Dan Buchal

The
[apps-script-portfolio-tracker](./packages/apps-script-portfolio-tracker)
package demonstrates how you could use the data from Empower to
automatically update the holdings, asset allocations, and other tables
in your copy of version 2.1.1 of Dan Buchal's very cool [Portfolio
Tracker Google
Sheet](https://github.com/danbuchal/portfolio-tracker). You will need
to enable the Google Sheets API service in the Services menu in Apps
Script.

### Minimal sample Apps Script code

```javascript
function doPost(e) {
  try {
    const {version: {major, minor}, holdings, classifications, accounts} = JSON.parse(e.postData.contents);
    if (major !== 0 || minor < 6) {
        throw new Error(`data version ${major}.${minor} not supported`);
    }
    Logger.log(`API version: ${major}.${minor}`);
    Logger.log(`Received ${holdings.length} holdings`);
    Logger.log(`Classifications for ${Object.keys(classifications).length} tickers`);
    Logger.log(`${accounts.length} accounts`);
    
    // Process data here (e.g., write to spreadsheet)
    
    return ContentService
      .createTextOutput(JSON.stringify({success: true, message: "Data received"}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

### Deployment Management

#### Initial deployment ####

1. In Apps Script editor, click "Deploy" → "New deployment".
2. Set type to "Web app".
3. Set "Execute as" to "Me".
4. Set "Who has access" to "Anyone".
5. Click "Deploy".
6. Copy the Web app URL for the deployment, and use it as the POST
   destination in the extension's configuration. You'll be prompted to
   allow the extension to access `https://script.google.com` URLs.

#### Deploying updates ####

To avoid having to change the URL in the extension's configuration, it
is advisable to deploy new code to an existing deployment as follows:

**Apps Script UI**
1. Select "Deploy" → "Manage deployments".
1. Select your deployment (typically it will be the only Active
   deployment) and click edit (pencil icon).
1. Change "Version" to "New version".
1. Click the "Deploy" button.

**Command line**

If you are using `[clasp](https://github.com/google/clasp?tab=readme-ov-file#clasp)` to manage your Apps Script code, you can push new code to the same deployment ID as follows:
```bash
clasp push
clasp create-version 'My Description'
clasp update-deployment --deploymentId YOUR_DEPLOYMENT_ID --version VERSION_NUMBER_FROM_CLASP_PUSH
```

## References

* This extension's [Chrome Web Store listing](https://chromewebstore.google.com/detail/empower-poster/lfjdkpiggkdkglapfjbifhgfhmilcmim).
* The [@righteffort/empower-poster-types](https://www.npmjs.com/package/@righteffort/empower-poster-types) package on npmjs.
* [Google guide to Apps Script Web Apps](https://developers.google.com/apps-script/guides/web).
* [clasp](https://github.com/google/clasp?tab=readme-ov-file#clasp).
* [Empower](https://ira.empower-retirement.com/).
