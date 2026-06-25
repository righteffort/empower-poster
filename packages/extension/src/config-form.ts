import { getPostUrls, normalizeUrl, setPostUrls } from "./util";

interface ConfigFormOptions {
  // Prefill the list with the currently-saved URLs (options page only).
  prefill: boolean;
  // Message shown after a successful save.
  savedMessage: string;
}

const EMPTY_WARNING =
  "No URLs configured — the extension will not do anything.";

export function initConfigForm({ prefill, savedMessage }: ConfigFormOptions) {
  const urlList = document.getElementById("urlList") as HTMLDivElement;
  const addButton = document.getElementById("addButton") as HTMLButtonElement;
  const saveButton = document.getElementById("saveButton") as HTMLButtonElement;
  const statusDiv = document.getElementById("status") as HTMLDivElement;

  function showStatus(message: string, type: string) {
    statusDiv.textContent = message;
    statusDiv.className = type;
  }

  function inputs() {
    return Array.from(urlList.querySelectorAll("input")) as HTMLInputElement[];
  }

  function warnIfEmpty() {
    if (inputs().length === 0) {
      showStatus(EMPTY_WARNING, "warning");
    }
  }

  function addRow(value = "") {
    const row = document.createElement("div");
    row.className = "url-row";

    const input = document.createElement("input");
    input.type = "url";
    input.placeholder = "https://api.example.com/exec";
    input.value = value;
    input.addEventListener("input", () => showStatus("", ""));

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.textContent = "Remove";
    removeButton.addEventListener("click", () => {
      row.remove();
      warnIfEmpty();
    });

    row.append(input, removeButton);
    urlList.append(row);
    return input;
  }

  addButton.addEventListener("click", () => addRow().focus());

  if (prefill) {
    void (async () => {
      const urls = await getPostUrls();
      urls.forEach((url) => addRow(url));
      warnIfEmpty();
    })();
  } else {
    addRow();
  }

  saveButton.addEventListener("click", async () => {
    const urls: string[] = [];
    for (const input of inputs()) {
      if (!input.value.trim()) continue;
      let url: string;
      try {
        url = normalizeUrl(input.value);
      } catch (e) {
        showStatus(e instanceof Error ? e.message : String(e), "error");
        input.focus();
        return;
      }
      if (urls.includes(url)) {
        showStatus(`Duplicate URL: ${url}`, "error");
        input.focus();
        return;
      }
      urls.push(url);
    }
    try {
      await setPostUrls(urls);
      showStatus(urls.length === 0 ? EMPTY_WARNING : savedMessage, "success");
    } catch (e) {
      const msg = `Failed: ${e instanceof Error ? e.message : e}`;
      showStatus(msg, "error");
      console.log(msg);
    }
  });
}
