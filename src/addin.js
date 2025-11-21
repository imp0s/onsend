const {
  base64ToUint8Array,
  removeMetadataAndComments,
  uint8ArrayToBase64,
} = require("./docCleanup");

async function getAttachmentContent(id) {
  return new Promise((resolve, reject) => {
    const item = getMailboxItem();
    item.getAttachmentContentAsync(id, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

async function removeAndReplaceAttachment(attachment) {
  const content = await getAttachmentContent(attachment.id);
  if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
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
        resolve();
      } else {
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
          resolve();
        } else {
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
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          const payload = arg;
          if (typeof payload.error === "number") {
            reject(new Error(`Dialog error: ${payload.error}`));
            return;
          }

          dialog.close();
          resolve(payload.message === "yes");
        });
      },
    );
  });
}

async function promptForCleanup(names) {
  const message = `Clean metadata and comments from: ${names.join(", ")}?`;
  if (typeof Office !== "undefined" && Office.context && Office.context.ui) {
    try {
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
    const item = getMailboxItem();
    const attachments = ((item && item.attachments) || []).filter(
      (att) => typeof att.isInline === "boolean",
    );
    const targets = attachments.filter((att) =>
      att.name.toLowerCase().endsWith(".docx"),
    );

    if (targets.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    const shouldClean = await promptForCleanup(targets.map((a) => a.name));
    if (!shouldClean) {
      event.completed({ allowEvent: true });
      return;
    }

    for (const attachment of targets) {
      await removeAndReplaceAttachment(attachment);
    }

    event.completed({ allowEvent: true });
  } catch (error) {
    console.error("Cleanup failed, sending without modification", error);
    event.completed({ allowEvent: true });
  }
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onMessageSend", onMessageSend);
}

module.exports = {
  onMessageSend,
  promptForCleanup,
  removeMetadataAndComments,
  base64ToUint8Array,
  uint8ArrayToBase64,
};
