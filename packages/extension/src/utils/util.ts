import { browser } from "#imports";

export async function getPostUrl(): Promise<string> {
  const val = (await browser.storage.sync.get("postUrl"))["postUrl"];
  return typeof val === "string" ? val : "";
}

export async function setPostUrl(postUrl: string) {
  const originPattern = getOriginPattern(postUrl);
  const granted = await browser.permissions.request({
    origins: [originPattern],
  });
  if (!granted) {
    throw new Error(
      `Permission was denied. Extension cannot POST data to .${postUrl}`,
    );
  }
  await browser.storage.sync.set({ postUrl });
}

function getOriginPattern(url: string) {
  try {
    const urlObj = new URL(url);
    return `${urlObj.origin}/*`;
  } catch {
    throw new Error(
      "Invalid URL format, should have form https://... or http://....",
    );
  }
}
