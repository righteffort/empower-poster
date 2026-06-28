import { getAccounts, getClassifications, getHoldings } from "./processing";
import type { ClassificationIn, HoldingEntryIn } from "./processing";
import type {
  Account,
  Classifications,
  HoldingEntry,
} from "@righteffort/empower-poster-types";
import type {
  PostDataRequest,
  PostDataResponse,
  TokenResponse,
  TokenUpdate,
  VisibilityUpdate,
} from "./types";

let csrf = "";

const fetchDataScript = async (csrf: string, path: string, params: object) => {
  try {
    const api_url = `https://pc-api.empower-retirement.com/${path}`;
    const body = new URLSearchParams({
      lastServerChangeId: "-1",
      csrf,
      apiClient: "WEB",
      ...params,
    }).toString();
    const options = {
      referrer: "https://ira.empower-retirement.com/",
      credentials: "include" as RequestCredentials,
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body,
    };
    const response = await fetch(api_url, options);
    if (!response.ok) {
      throw new Error(`${api_url} response status: ${response.status}`);
    }
    const data = await response.json();
    return data;
  } catch (e) {
    const msg = `Caught exception: ${e instanceof Error ? e.message : e}`;
    console.error(msg);
    return { error: msg };
  }
};

const container = document.createElement("div");
container.style.display = "none";

const button = document.createElement("button");
button.textContent = "Post data";
button.style.cssText =
  "position: fixed; top: 10px; right: 10px; z-index: 10000; padding: 10px; background: #007bff; color: white; border: none; cursor: pointer; border-radius: 4px;";

const statusLine = document.createElement("div");
statusLine.id = "post-data-status";
statusLine.style.cssText =
  "position: fixed; top: 60px; right: 10px; z-index: 10000; padding: 5px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px; font-size: 12px; max-width: 200px;";
statusLine.textContent = "Ready";

container.appendChild(button);
container.appendChild(statusLine);

button.onclick = async () => {
  button.disabled = true;
  statusLine.textContent = "Posting...";
  try {
    if (!csrf) {
      throw new Error("Can't retrieve data. Not logged in?");
    }
    const getHoldingsRawData = (
      await fetchDataScript(csrf, "api/invest/getHoldings", {
        userAccountIds: "[]",
        classificationStyles: '["allocation"]',
      })
    ).spData;
    const classificationsIn: ClassificationIn[] =
      getHoldingsRawData?.classifications?.[0]?.classifications;
    if (classificationsIn == null) {
      throw new Error(
        "spData.classifications is missing required nested classifications holdings JSON, nothing to POST",
      );
    }
    const holdingsIn: HoldingEntryIn[] = getHoldingsRawData.holdings;
    if (holdingsIn == null) {
      throw new Error(
        "spData.holdings missing from holdings JSON data, nothing to POST",
      );
    }
    const accountsIn = (
      await fetchDataScript(csrf, "api/newaccount/getAccounts2", {
        userAccountIds: "[]",
        classificationStyles: '["allocation"]',
      })
    ).spData?.accounts;
    if (accountsIn == null) {
      throw new Error(
        "spData.accounts missing from accounts JSON data, nothing to POST",
      );
    }
    const accounts = getAccounts(accountsIn);
    const holdings = getHoldings(holdingsIn);
    const { classifications, errors: classificationsErrors } =
      getClassifications(classificationsIn);

    if (classificationsErrors.length > 0) {
      if (classificationsErrors.length > 20) {
        statusLine.innerHTML = `${classificationsErrors.length} negative categorizations found, see console for details`;
        console.log("Negative categorizations:", classificationsErrors);
      } else {
        const errorList = classificationsErrors
          .map((error) => {
            const pct = parseFloat((error.fraction * 100).toFixed(6));
            const classStr = error.classes.filter(Boolean).join(":");
            return `${error.ticker} ${classStr} ${pct}%`;
          })
          .join("<br>");
        statusLine.innerHTML = `Negative categorization(s):<br>${errorList}`;
      }

      // Create buttons for user decision
      const buttonContainer = document.createElement("div");
      buttonContainer.style.cssText = "margin-top: 10px;";

      const cancelBtn = document.createElement("button");
      cancelBtn.textContent = "Cancel";
      cancelBtn.style.cssText =
        "width: 100%; box-sizing: border-box; margin-bottom: 5px; padding: 5px 10px; background: #007bff; color: white; border: none; cursor: pointer; border-radius: 4px;";
      cancelBtn.focus();

      const postBtn = document.createElement("button");
      postBtn.textContent = "Post anyway, with negative amounts zeroed";
      postBtn.style.cssText =
        "width: 100%; box-sizing: border-box; padding: 5px 10px; background: #dc3545; color: white; border: none; cursor: pointer; border-radius: 4px;";

      cancelBtn.onclick = () => {
        statusLine.textContent = "Ready";
        statusLine.innerHTML = "";
        button.disabled = false;
        button.textContent = "Post data";
      };

      postBtn.onclick = async () => {
        buttonContainer.remove();
        statusLine.textContent = "Posting...";
        await postProcessedData(holdings, classifications, accounts);
      };

      buttonContainer.appendChild(cancelBtn);
      buttonContainer.appendChild(postBtn);
      statusLine.appendChild(buttonContainer);

      button.disabled = false;
      button.textContent = "Post data";
      return;
    }

    await postProcessedData(holdings, classifications, accounts);
  } catch (e) {
    const msg = e instanceof Error ? e.message : JSON.stringify(e);
    statusLine.textContent = msg;
    console.error(msg);
    button.disabled = false;
    button.textContent = "Post data";
  }
};

async function postProcessedData(
  holdings: HoldingEntry[],
  classifications: Classifications,
  accounts: Account[],
) {
  try {
    const message: PostDataRequest = {
      type: "POST_DATA_REQUEST",
      data: { holdings, classifications, accounts },
    };
    const response = (await chrome.runtime.sendMessage(
      message,
    )) as PostDataResponse;
    if (!response.ok) {
      throw new Error(
        `POST failed: ${response.message}. Do you need to configure POST destination at chrome://extensions/?options=lfjdkpiggkdkglapfjbifhgfhmilcmim ?`,
      );
    }
    statusLine.textContent = `Posted at ${new Date()}`;
  } catch (e) {
    const msg = e instanceof Error ? e.message : JSON.stringify(e);
    statusLine.textContent = msg;
    console.error(msg);
  } finally {
    button.disabled = false;
    button.textContent = "Post data";
  }
}

// Listen for token updates and visibility changes from background script
chrome.runtime.onMessage.addListener(
  (message: TokenUpdate | VisibilityUpdate) => {
    if (message.type === "TOKEN_UPDATE" && message.csrf && !csrf) {
      csrf = message.csrf;
    }
    if (message.type === "VISIBILITY_UPDATE" && message.show !== undefined) {
      container.style.display = message.show ? "block" : "none";
    }
  },
);

// Fallback token request after 5 seconds
setTimeout(async () => {
  if (!csrf) {
    try {
      const response = (await chrome.runtime.sendMessage({
        type: "TOKEN_REQUEST",
      })) as TokenResponse;

      if (response.csrf) {
        csrf = response.csrf;
      }
    } catch (e) {
      console.log(
        `token request failed: ${e instanceof Error ? e.message : e}`,
      );
    }
  }
}, 5000);

container.style.display = window.location.pathname.startsWith("/dashboard/")
  ? "block"
  : "none";
document.body.appendChild(container);
