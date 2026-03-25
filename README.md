# Empower Poster

Browser extension that uploads holdings and asset classification data
from [Empower](https://ira.empower-retirement.com/) (formerly Personal
Capital) to a user-specified external HTTPS endpoint. For example, the
endpoint could be an Apps Script web app bound to a Google Sheets
spreadsheet, and could update the spreadsheet based on the uploaded
data.

## Endpoint request/response API

The types referenced below are defined in [`api.ts`](./packages/extension/src/api.ts).

### Request

The extension makes HTTPS POST requests with `Content-Type:
application/json`. The payload conforms to `PostPayload` and contains:
- `version`: API version.
- `holdings`: Array of `HoldingEntry` objects with account, ticker,
  price, quantity, and value.
- `classifications`: `Classifications` object mapping tickers to asset
  classifications.
  - For each ticker there is one or more Classification entries, with
  fractions that add to 1.
  - Each Classification has the asset class, asset subclass (possibly
  empty), and fraction.

### Response

The response should have `Content-Type: application/json`, and the
body should conform to one of the response types:
- `SuccessResponse`: `{success: true, message?: string}`.
- `ErrorResponse`: `{success: false, error: string}`.

**Response Handling:**

The extension interprets the response to the POST request as follows:
- Non-2xx HTTPS status codes are treated as failures.
- Content-Type `application/json`: success based on the `success` field in the JSON response body.
- Other values of Content-Types: empty body indicates success, non-empty is a failure message.

## Using in Google Sheets Apps Script

You can start by making a copy of [this spreadsheet](https://docs.google.com/spreadsheets/d/1UlLLp5KLMbtYUJEV35B92Rj7j8ivbopPp-INh4W8lYw/edit)

### Apps Script Template

```javascript
function doPost(e) {
  try {
    const {version, holdings, classifications} = JSON.parse(e.postData.contents);

    console.log(`Received ${holdings.length} holdings`);
    console.log(`Classifications for ${Object.keys(classifications).length} tickers`);
    console.log(`API version: ${version}`);
    
	if (version !== "0.3") {
		throw new Error(`data version ${version} not supported`);
	}
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

* [Google guide to Apps Script Web Apps](https://developers.google.com/apps-script/guides/web).
* [clasp](https://github.com/google/clasp?tab=readme-ov-file#clasp).
* [Empower](https://ira.empower-retirement.com/).
* This extension's [Chrome Web Store listing](https://chromewebstore.google.com/detail/empower-poster/lfjdkpiggkdkglapfjbifhgfhmilcmim).
