import type { PostPayload } from "@righteffort/empower-poster-types";
import type {
  PostDataRequest,
  PostDataResponse,
  TokenRequest,
  TokenResponse,
} from "./types";
import { getPostUrls } from "./util";

console.log(`Background script started at ${new Date()}`);

chrome.runtime.onInstalled.addListener(async (details) => {
  console.log(`Extension installed at ${new Date()}`, details);
  if ((await getPostUrls()).length === 0) {
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
    const postUrls = await getPostUrls();
    if (postUrls.length === 0) {
      throw new Error("No POST URLs configured, can't post data");
    }
    if (!data.holdings || !data.classifications || !data.accounts) {
      console.log(`data=${JSON.stringify(data)}`);
      throw new Error(
        "holdings, classifications, or accounts missing from processed data",
      );
    }
    const payload: PostPayload = {
      version: { major: 0, minor: 5 },
      holdings: data.holdings,
      classifications: data.classifications,
      accounts: data.accounts,
    };
    const options = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    };
    const results = await Promise.all(
      postUrls.map((url) => postToUrl(url, options)),
    );
    const ok = results.every((r) => r.ok);
    const message = results.map((r) => `${r.url}: ${r.message}`).join("\n");
    return { ok, message };
  } catch (e) {
    return {
      ok: false,
      message: e instanceof Error ? e.message : JSON.stringify(e),
    };
  }
}

async function postToUrl(
  url: string,
  options: RequestInit,
): Promise<{ url: string; ok: boolean; message: string }> {
  try {
    const r = await fetch(url, options);
    if (!r.ok) {
      return { url, ok: false, message: `${r.status}: ${await r.text()}` };
    }
    const contentType = r.headers.get("content-type");
    if (contentType?.includes("application/json")) {
      const json = await r.json();
      const ok = json.success === true && !json.error;
      const message = json.message || json.error || JSON.stringify(json);
      return { url, ok, message };
    }
    const detail = await r.text();
    return { url, ok: !detail, message: detail };
  } catch (e) {
    return {
      url,
      ok: false,
      message: e instanceof Error ? e.message : JSON.stringify(e),
    };
  }
}
