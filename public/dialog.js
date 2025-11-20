/* global Office */

(function () {
  const messageElement = document.getElementById("message");
  const messageFromHash = decodeURIComponent(
    window.location.hash.replace("#", ""),
  );
  messageElement.textContent = messageFromHash || "Clean attachments?";

  document.getElementById("yes")?.addEventListener("click", () => {
    Office.context.ui.messageParent("yes");
  });

  document.getElementById("no")?.addEventListener("click", () => {
    Office.context.ui.messageParent("no");
  });
})();
