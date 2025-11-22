const {
  base64ToUint8Array,
  removeMetadataAndComments,
  uint8ArrayToBase64,
} = require("./docCleanup");

async function getAttachmentContent(id) {
  return new Promise((resolve, reject) => {
    const item = getMailboxItem();
    console.log("[addin] fetching attachment content", id);
    item.getAttachmentContentAsync(id, (result) => {
      if (
        result.status === Office.AsyncResultStatus.Succeeded &&
        result.value
      ) {
        console.log("[addin] fetched attachment content", {
          id,
          format: result.value.format,
        });
        resolve(result.value);
      } else {
        console.error("[addin] failed to fetch attachment", id, result.error);
        reject(result.error);
      }
    });
  });
}

async function removeAndReplaceAttachment(attachment) {
  console.log("[addin] cleaning attachment", {
    id: attachment.id,
    name: attachment.name,
    isInline: attachment.isInline,
  });
  const content = await getAttachmentContent(attachment.id);
  if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
    console.error("[addin] unsupported attachment format", {
      id: attachment.id,
      format: content.format,
    });
    throw new Error("Unsupported attachment format");
  }

  const cleaned = await removeMetadataAndComments(
    base64ToUint8Array(content.content),
  );
  const updatedBase64 = uint8ArrayToBase64(cleaned);

  await new Promise((resolve, reject) => {
    const item = getMailboxItem();
    item.removeAttachmentAsync(attachment.id, (removeResult) => {
      if (removeResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("[addin] removed attachment", attachment.id);
        resolve();
      } else {
        console.error(
          "[addin] failed to remove attachment",
          attachment.id,
          removeResult.error,
        );
        reject(removeResult.error);
      }
    });
  });

  await new Promise((resolve, reject) => {
    const item = getMailboxItem();
    item.addFileAttachmentFromBase64Async(
      updatedBase64,
      attachment.name,
      { isInline: attachment.isInline },
      (addResult) => {
        if (addResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("[addin] re-attached cleaned file", {
            id: attachment.id,
            name: attachment.name,
          });
          resolve();
        } else {
          console.error(
            "[addin] failed to add cleaned attachment",
            attachment.id,
            addResult.error,
          );
          reject(addResult.error);
        }
      },
    );
  });
}

function getMailboxItem() {
  const item = Office.context?.mailbox?.item;
  if (!item) {
    throw new Error("Mailbox item unavailable");
  }
  return item;
}

function openConfirmationDialog(message) {
  return new Promise((resolve, reject) => {
    const url = `${window.location.origin}/public/dialog.html#${encodeURIComponent(
      message,
    )}`;
    Office.context.ui.displayDialogAsync(
      url,
      { height: 40, width: 30 },
      (result) => {
        if (
          result.status !== Office.AsyncResultStatus.Succeeded ||
          !result.value
        ) {
          reject(result.error);
          return;
        }

        const dialog = result.value;
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg) => {
            const payload = arg;
            if (typeof payload.error === "number") {
              reject(new Error(`Dialog error: ${payload.error}`));
              return;
            }

            dialog.close();
            resolve(payload.message === "yes");
          },
        );
      },
    );
  });
}

async function promptForCleanup(names) {
  const message = `Clean metadata and comments from: ${names.join(", ")}?`;
  if (typeof Office !== "undefined" && Office.context && Office.context.ui) {
    try {
      console.log("[addin] opening confirmation dialog for attachments", names);
      return await openConfirmationDialog(message);
    } catch (error) {
      console.error("Dialog failed; defaulting to proceed", error);
      return true;
    }
  }

  return typeof window !== "undefined" ? window.confirm(message) : true;
}

async function onMessageSend(event) {
  try {
    console.log("[addin] onMessageSend invoked");
    const item = getMailboxItem();
    const attachments = ((item && item.attachments) || []).filter(
      (att) => typeof att.isInline === "boolean",
    );
    const targets = attachments.filter((att) =>
      att.name.toLowerCase().endsWith(".docx"),
    );

    console.log("[addin] attachments discovered", {
      total: attachments.length,
      candidates: targets.length,
    });

    if (targets.length === 0) {
      console.log("[addin] no Word attachments detected; allowing send");
      event.completed({ allowEvent: true });
      return;
    }

    const shouldClean = await promptForCleanup(targets.map((a) => a.name));
    if (!shouldClean) {
      console.log("[addin] user declined cleanup; allowing send");
      event.completed({ allowEvent: true });
      return;
    }

    for (const attachment of targets) {
      console.log("[addin] processing attachment", attachment.name);
      await removeAndReplaceAttachment(attachment);
    }

    console.log("[addin] cleanup complete; allowing send");
    event.completed({ allowEvent: true });
  } catch (error) {
    console.error("Cleanup failed, sending without modification", error);
    event.completed({ allowEvent: true });
  }
}

console.log("[addin] script loaded; awaiting Office.onReady");
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => {
    console.log("[addin] Office.ready resolved; wiring handlers");
    if (Office.actions) {
      console.log("[addin] associating onMessageSend action");
      Office.actions.associate("onMessageSend", onMessageSend);
    }

    if (typeof window !== "undefined") {
      window.onMessageSend = onMessageSend;
    }
  });
}

module.exports = {
  onMessageSend,
  promptForCleanup,
  removeMetadataAndComments,
  base64ToUint8Array,
  uint8ArrayToBase64,
};
