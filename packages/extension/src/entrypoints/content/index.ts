import { getClassifications, getHoldings } from "./processing";
import type {
  Classifications,
  HoldingEntry,
} from "@righteffort/empower-poster-types";
import type {
  PostDataRequest,
  PostDataResponse,
  TokenResponse,
} from "../../utils/types";

export default defineContentScript({
  matches: [
    "https://ira.empower-retirement.com/*",
    "https://participant.empower-retirement.com/*",
  ],
  main() {
    let csrf = "";

    const fetchDataScript = async (csrf: string) => {
      try {
        const api_url =
          "https://pc-api.empower-retirement.com/api/invest/getHoldings";
        const body = new URLSearchParams({
          userAccountIds: "[]",
          classificationStyles: '["allocation"]',
          lastServerChangeId: "-1",
          csrf,
          apiClient: "WEB",
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
        const resultsJson = await fetchDataScript(csrf);

        if (!resultsJson.spData) {
          throw new Error("spData missing from JSON data, nothing to POST");
        }

        const classificationsIn =
          resultsJson.spData.classifications[0].classifications;
        const holdings = getHoldings(resultsJson.spData.holdings);
        const classificationsResult = getClassifications(classificationsIn);

        if (classificationsResult.errors.length > 0) {
          if (classificationsResult.errors.length > 20) {
            statusLine.innerHTML = `${classificationsResult.errors.length} negative categorizations found, see console for details`;
            console.log(
              "Negative categorizations:",
              classificationsResult.errors,
            );
          } else {
            const errorList = classificationsResult.errors
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
            await postProcessedData(
              holdings,
              classificationsResult.classifications,
            );
          };

          buttonContainer.appendChild(cancelBtn);
          buttonContainer.appendChild(postBtn);
          statusLine.appendChild(buttonContainer);

          button.disabled = false;
          button.textContent = "Post data";
          return;
        }

        await postProcessedData(
          holdings,
          classificationsResult.classifications,
        );
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
    ) {
      try {
        const message: PostDataRequest = {
          type: "POST_DATA_REQUEST",
          data: { holdings, classifications },
        };
        const response = (await browser.runtime.sendMessage(
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

    // Listen for token updates from background script
    browser.runtime.onMessage.addListener((message) => {
      if (message.type === "TOKEN_UPDATE" && message.csrf && !csrf) {
        csrf = message.csrf;
        container.style.display = "block";
      }
    });

    // Fallback token request after 5 seconds
    setTimeout(async () => {
      if (!csrf) {
        try {
          const response = (await browser.runtime.sendMessage({
            type: "TOKEN_REQUEST",
          })) as TokenResponse;

          if (response.csrf) {
            csrf = response.csrf;
            container.style.display = "block";
          }
        } catch (e) {
          console.log(
            `token request failed: ${e instanceof Error ? e.message : e}`,
          );
        }
      }
    }, 5000);

    document.body.appendChild(container);
  },
});
