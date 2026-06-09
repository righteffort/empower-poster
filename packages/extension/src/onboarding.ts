// TODO: merge {onboarding,options}.{html,ts}.

import { setPostUrl } from "./util";

const urlInput = document.getElementById("postUrl") as HTMLInputElement;
const saveButton = document.getElementById("saveButton") as HTMLButtonElement;
const statusDiv = document.getElementById("status") as HTMLDivElement;

function showStatus(message: string, type: string) {
  statusDiv.textContent = message;
  statusDiv.className = type;
}

urlInput.addEventListener("input", () => showStatus("", ""));

saveButton.addEventListener("click", async () => {
  const postUrl = urlInput.value.trim();
  if (!postUrl) {
    showStatus("Please enter a URL.", "error");
    return;
  }
  try {
    await setPostUrl(postUrl);
    showStatus("Settings saved, you may close this window.", "success");
    // setTimeout(() => {
    //  window.close();
    // }, 3000);
  } catch (e) {
    const msg = `Failed: ${e instanceof Error ? e.message : e}`;
    showStatus(msg, "error");
    console.log(msg);
  }
});
