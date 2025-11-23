const BLOCK_MESSAGE =
  "Cannot send message as it has attachments we cannot verify. Use your desktop client.";

function getMailboxItem() {
  const item = Office.context?.mailbox?.item;
  if (!item) {
    throw new Error("Mailbox item unavailable");
  }
  return item;
}

function getAttachments() {
  const item = getMailboxItem();

  if (typeof item.getAttachmentsAsync !== "function") {
    const cached = Array.isArray(item.attachments) ? item.attachments : [];
    console.log("[addin] attachments pulled from cached item state", {
      total: cached.length,
    });
    return Promise.resolve(cached);
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

function showBlockingNotification() {
  const item = getMailboxItem();
  return new Promise((resolve, reject) => {
    item.notificationMessages.addAsync(
      "AttachmentBlock",
      {
        type: Office.MailboxEnums.ItemNotificationMessageType
          .InformationalMessage,
        message: BLOCK_MESSAGE,
        icon: "icon32",
        persistent: true,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("[addin] displayed attachment block notification");
          resolve();
        } else {
          console.error("[addin] failed to display notification", result.error);
          reject(result.error);
        }
      },
    );
  });
}

async function onMessageSend(event) {
  try {
    console.log("[addin] onMessageSend invoked");
    const attachments = await getAttachments();

    if (!attachments || attachments.length === 0) {
      console.log("[addin] no attachments detected; allowing send");
      event.completed({ allowEvent: true });
      return;
    }

    console.log("[addin] attachments found; blocking send", {
      total: attachments.length,
    });

    try {
      await showBlockingNotification();
    } catch (error) {
      console.error("[addin] notification failed", error);
    }

    event.completed({ allowEvent: false });
  } catch (error) {
    console.error("[addin] unexpected error; allowing send", error);
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
  getAttachments,
  showBlockingNotification,
};
