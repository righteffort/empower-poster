// TODO: merge {onboarding,options}.{html,ts}.

import { getPostUrl, setPostUrl } from "./util";

const urlInput = document.getElementById("postUrl") as HTMLInputElement;
const saveButton = document.getElementById("saveButton") as HTMLButtonElement;
const statusDiv = document.getElementById("status") as HTMLDivElement;

function showStatus(message: string, type: string) {
  statusDiv.textContent = message;
  statusDiv.className = type;
}

urlInput.addEventListener("input", () => showStatus("", ""));

await (async function () {
  const postUrl = await getPostUrl();
  if (postUrl) {
    urlInput.value = postUrl;
  }
})();

saveButton.addEventListener("click", async () => {
  const postUrl = urlInput.value.trim();
  if (!postUrl) {
    showStatus("Please enter a URL.", "error");
    return;
  }
  try {
    await setPostUrl(postUrl);
    showStatus("Settings saved", "success");
  } catch (e) {
    const msg = `Failed: ${e instanceof Error ? e.message : e}`;
    showStatus(msg, "error");
    console.log(msg);
  }
});
