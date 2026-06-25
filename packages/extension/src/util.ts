export async function getPostUrls(): Promise<string[]> {
  const stored = await chrome.storage.sync.get(["postUrls", "postUrl"]);
  if (Array.isArray(stored["postUrls"])) {
    return stored["postUrls"].filter((u): u is string => typeof u === "string");
  }
  // Transparently upgrade a legacy single `postUrl` to `postUrls`.
  const legacy = stored["postUrl"];
  if (typeof legacy === "string" && legacy) {
    const postUrls = [legacy];
    await chrome.storage.sync.set({ postUrls });
    await chrome.storage.sync.remove("postUrl");
    return postUrls;
  }
  return [];
}

export async function setPostUrls(postUrls: string[]) {
  if (postUrls.length > 0) {
    const granted = await chrome.permissions.request({
      origins: postUrls.map(getOriginPattern),
    });
    if (!granted) {
      throw new Error(
        "Permission was denied. Extension cannot POST data to one or more of the configured URLs.",
      );
    }
  }
  await chrome.storage.sync.set({ postUrls });
  // Drop the legacy key on every write so it never resurfaces.
  await chrome.storage.sync.remove("postUrl");
}

// Parse a (trimmed) URL, throwing a friendly error on invalid input.
function parseUrl(raw: string): URL {
  try {
    return new URL(raw.trim());
  } catch {
    throw new Error(
      "Invalid URL format, should have form https://... or http://....",
    );
  }
}

// Canonicalize a URL: the URL parser lowercases the scheme and hostname while
// preserving the case of the path, query, and fragment. Throws on invalid input.
export function normalizeUrl(raw: string): string {
  return parseUrl(raw).href;
}

function getOriginPattern(url: string) {
  return `${parseUrl(url).origin}/*`;
}
