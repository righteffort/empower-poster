import type { PostPayload } from "@righteffort/empower-poster-types";
import type {
  PostDataRequest,
  PostDataResponse,
  TokenRequest,
  TokenResponse,
} from "./types";
import { getPostUrl } from "./util";

console.log(`Background script started at ${new Date()}`);

chrome.runtime.onInstalled.addListener(async (details) => {
  console.log(`Extension installed at ${new Date()}`, details);
  if (!(await getPostUrl())) {
    chrome.tabs.create({ url: "src/onboarding.html" });
  }
});

let csrf = "";

chrome.webRequest.onBeforeRequest.addListener(
  (details) => {
    const csrfVal = details?.requestBody?.formData?.["csrf"]?.[0];
    if (csrfVal && csrfVal !== csrf) {
      csrf = csrfVal as string;
      console.log("Saved token");
      // Send CSRF token to all tabs with content script
      chrome.tabs.query({}).then((tabs) => {
        for (const tab of tabs) {
          if (tab.id) {
            chrome.tabs
              .sendMessage(tab.id, {
                type: "TOKEN_UPDATE",
                csrf,
              })
              .catch(() => {
                // Ignore tabs that don't have content script loaded
              });
          }
        }
      });
    }
    return {};
  },
  {
    urls: ["https://pc-api.empower-retirement.com/*"],
    types: ["xmlhttprequest"],
  },
  ["requestBody"],
);

// Listen for post data message from content script
chrome.runtime.onMessage.addListener(
  (message: TokenRequest | PostDataRequest) => {
    // TODO: try catch, and return a response either way.
    if (message.type === "POST_DATA_REQUEST") {
      const request = message as PostDataRequest;
      return postData(request.data);
    }
    if (message.type === "TOKEN_REQUEST") {
      const response: TokenResponse = { csrf };
      return Promise.resolve(response);
    }
    return;
  },
);

async function postData(
  data: any, // eslint-disable-line @typescript-eslint/no-explicit-any
): Promise<PostDataResponse> {
  try {
    const postUrl = await getPostUrl();
    if (!postUrl) {
      throw new Error("POST URL not configured, can't post data");
    }
    if (!data.holdings || !data.classifications) {
      console.log(`data=${JSON.stringify(data)}`);
      throw new Error(
        "holdings or classifications missing from processed data",
      );
    }
    const payload: PostPayload = {
      version: "0.3",
      holdings: data.holdings,
      classifications: data.classifications,
    };
    const options = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    };
    const r = await fetch(postUrl, options);
    if (!r.ok) {
      return { ok: false, message: `${r.status}: ${await r.text()}` };
    }
    const contentType = r.headers.get("content-type");
    let detail: string;
    let ok: boolean;
    if (contentType?.includes("application/json")) {
      const json = await r.json();
      ok = json.success === true && !json.error;
      detail = json.message || json.error || JSON.stringify(json);
    } else {
      detail = await r.text();
      ok = !detail;
    }
    return { ok, message: detail };
  } catch (e) {
    return {
      ok: false,
      message: e instanceof Error ? e.message : JSON.stringify(e),
    };
  }
}
