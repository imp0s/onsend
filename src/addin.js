const {
  base64ToUint8Array,
  removeMetadataAndComments,
  uint8ArrayToBase64,
} = require("./docCleanup");

function normalizeAttachmentId(attachment) {
  if (attachment.id) return attachment.id;
  if (attachment.attachmentId) return attachment.attachmentId;
  throw new Error("Attachment identifier missing");
}

async function getAttachmentContent(id) {
  const item = getMailboxItem();
  const ewsContent = await tryGetAttachmentContentFromEws(id);
  if (ewsContent) {
    return ewsContent;
  }

  console.log("[addin] fetching attachment content", id);

  const result = await new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(id, (asyncResult) => {
      if (
        asyncResult.status === Office.AsyncResultStatus.Succeeded &&
        asyncResult.value
      ) {
        resolve(asyncResult.value);
      } else {
        reject(asyncResult.error);
      }
    });
  });

  if (result.format === Office.MailboxEnums.AttachmentContentFormat.FileUrl) {
    const fallback = await tryGetAttachmentContentFromEws(id);
    if (fallback) {
      return fallback;
    }

    console.warn("[addin] falling back to fileUrl content fetch", { id });
    return result;
  }

  if (result.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    console.log("[addin] fetched attachment content", {
      id,
      format: result.format,
    });
    return result;
  }

  console.error("[addin] unsupported attachment content format", {
    id,
    format: result.format,
  });
  throw new Error("Unsupported attachment format");
}

async function tryGetAttachmentContentFromEws(id) {
  const mailbox = Office.context?.mailbox;
  if (!mailbox || typeof mailbox.makeEwsRequestAsync !== "function") {
    return null;
  }

  const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2016" />
  </soap:Header>
  <soap:Body>
    <m:GetAttachment>
      <m:AttachmentShape>
        <t:IncludeMimeContent>true</t:IncludeMimeContent>
      </m:AttachmentShape>
      <m:AttachmentIds>
        <t:AttachmentId Id="${id}" />
      </m:AttachmentIds>
    </m:GetAttachment>
  </soap:Body>
</soap:Envelope>`;

  const ewsResponse = await new Promise((resolve, reject) => {
    mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult.value);
      } else {
        reject(
          asyncResult.error || new Error("Failed to fetch attachment via EWS"),
        );
      }
    });
  });

  const match = ewsResponse.match(/<t:Content>([\s\S]*?)<\/t:Content>/i);
  if (!match || !match[1]) {
    console.error("[addin] EWS response missing content", { id });
    return null;
  }

  console.log("[addin] fetched attachment via EWS", { id });
  return {
    format: Office.MailboxEnums.AttachmentContentFormat.Base64,
    content: match[1],
  };
}

async function removeAndReplaceAttachment(attachment) {
  console.log("[addin] cleaning attachment", {
    id: normalizeAttachmentId(attachment),
    name: attachment.name,
    isInline: attachment.isInline,
  });
  const content = await getAttachmentContent(normalizeAttachmentId(attachment));
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
    item.removeAttachmentAsync(
      normalizeAttachmentId(attachment),
      (removeResult) => {
        if (removeResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log(
            "[addin] removed attachment",
            normalizeAttachmentId(attachment),
          );
          resolve();
        } else {
          console.error(
            "[addin] failed to remove attachment",
            normalizeAttachmentId(attachment),
            removeResult.error,
          );
          reject(removeResult.error);
        }
      },
    );
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

async function getAttachments() {
  const item = getMailboxItem();
  if (typeof item.getAttachmentsAsync !== "function") {
    const cached = Array.isArray(item.attachments) ? item.attachments : [];
    console.log("[addin] attachments pulled from cached item state", {
      total: cached.length,
    });
    return cached;
  }

  return new Promise((resolve, reject) => {
    item.getAttachmentsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const list = Array.isArray(result.value) ? result.value : [];
        console.log("[addin] attachments fetched via getAttachmentsAsync", {
          total: list.length,
        });
        resolve(list);
      } else {
        console.error("[addin] failed to fetch attachments", result.error);
        reject(result.error);
      }
    });
  });
}

async function onMessageSend(event) {
  try {
    console.log("[addin] onMessageSend invoked");
    const attachments = (await getAttachments()).filter(
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
  getAttachments,
  getAttachmentContent,
  base64ToUint8Array,
  uint8ArrayToBase64,
  normalizeAttachmentId,
};
