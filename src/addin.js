const { allowedDomainExtensions } = require("./config");

const ATTACHMENT_BLOCK_MESSAGE =
  "Cannot send messages with attachments. Remove all attachments and try again.";
const RECIPIENT_BLOCK_MESSAGE =
  "Message blocked: recipients must use the same domain as the sender.";
const CONFIG_BLOCK_MESSAGE =
  "Message blocked: configure allowed domains before sending.";
const BLOCK_NOTIFICATION_ID = "SafetyCheckBlock";

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

function extractDomain(email) {
  if (typeof email !== "string") return null;
  const atIndex = email.lastIndexOf("@");
  if (atIndex < 0 || atIndex === email.length - 1) return null;
  return email.slice(atIndex + 1).toLowerCase();
}

function normalizeRecipientAddress(recipient) {
  if (!recipient) return null;
  if (typeof recipient === "string") return recipient;
  if (typeof recipient.emailAddress === "string") return recipient.emailAddress;
  if (typeof recipient.Address === "string") return recipient.Address;
  return null;
}

function getAllowedDomains() {
  const senderDomain = extractDomain(
    Office.context?.mailbox?.userProfile?.emailAddress,
  );
  if (senderDomain) {
    return { domains: [senderDomain], enforceExact: true };
  }

  const configured = Array.isArray(allowedDomainExtensions)
    ? allowedDomainExtensions
    : [];
  const domains = configured
    .map((domain) => (typeof domain === "string" ? domain.trim() : ""))
    .filter(Boolean)
    .map((domain) => domain.toLowerCase());

  return { domains, enforceExact: false };
}

function isDomainAllowed(domain, allowedDomains, enforceExact) {
  if (!domain || !allowedDomains || allowedDomains.length === 0) return false;

  if (enforceExact) {
    return allowedDomains.includes(domain.toLowerCase());
  }

  return allowedDomains.some((allowed) => {
    const candidate = allowed.toLowerCase();
    return (
      domain.toLowerCase() === candidate ||
      domain.toLowerCase().endsWith(`.${candidate}`) ||
      domain.toLowerCase().endsWith(candidate)
    );
  });
}

function recipientsWithDisallowedDomains(recipients, allowedDomains, enforceExact) {
  if (!Array.isArray(recipients)) return [];

  return recipients
    .map(normalizeRecipientAddress)
    .map(extractDomain)
    .map((domain, index) => ({ domain, original: recipients[index] }))
    .filter((entry) => !isDomainAllowed(entry.domain, allowedDomains, enforceExact));
}

function getRecipients(fieldName) {
  const item = getMailboxItem();
  const field = item?.[fieldName];

  if (field && typeof field.getAsync === "function") {
    return new Promise((resolve, reject) => {
      field.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(Array.isArray(result.value) ? result.value : []);
        } else {
          reject(result.error);
        }
      });
    });
  }

  const cached = Array.isArray(field) ? field : [];
  return Promise.resolve(cached);
}

async function getAllRecipients() {
  const [to, cc, bcc] = await Promise.all([
    getRecipients("to"),
    getRecipients("cc"),
    getRecipients("bcc"),
  ]);

  return [...to, ...cc, ...bcc];
}

function showBlockNotification(item, message) {
  const notificationType =
    Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;
  const notificationPayload = {
    type: notificationType,
    message,
    icon: "icon32",
    persistent: true,
  };

  return new Promise((resolve) => {
    item.notificationMessages.replaceAsync(
      BLOCK_NOTIFICATION_ID,
      notificationPayload,
      () => resolve(),
    );
  });
}

async function validateRecipientDomains() {
  const { domains: allowedDomains, enforceExact } = getAllowedDomains();
  if (allowedDomains.length === 0) {
    return { allowed: false, reason: "no-allowed-domains", offenders: [] };
  }

  const recipients = await getAllRecipients();
  const offenders = recipientsWithDisallowedDomains(
    recipients,
    allowedDomains,
    enforceExact,
  );

  if (offenders.length > 0) {
    return { allowed: false, reason: "domain-mismatch", offenders };
  }

  return { allowed: true, reason: null, offenders: [] };
}

async function onMessageSend(event) {
  try {
    console.log("[addin] onMessageSend invoked");

    const [attachments, domainCheck] = await Promise.all([
      getAttachments(),
      validateRecipientDomains(),
    ]);

    const item = getMailboxItem();

    if (!domainCheck.allowed) {
      const message =
        domainCheck.reason === "no-allowed-domains"
          ? CONFIG_BLOCK_MESSAGE
          : RECIPIENT_BLOCK_MESSAGE;
      console.log("[addin] blocking send due to recipient domains", {
        reason: domainCheck.reason,
        offenders: domainCheck.offenders.map((offender) => offender.domain),
      });
      await showBlockNotification(item, message);
      event.completed({ allowEvent: false });
      return;
    }

    if (attachments && attachments.length > 0) {
      console.log("[addin] attachments found; blocking send", {
        total: attachments.length,
      });
      await showBlockNotification(item, ATTACHMENT_BLOCK_MESSAGE);
      event.completed({ allowEvent: false });
      return;
    }

    console.log("[addin] safety checks passed; allowing send");
    event.completed({ allowEvent: true });
  } catch (error) {
    console.error("[addin] unexpected error during safety checks", error);
    event.completed({ allowEvent: false });
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
  extractDomain,
  isDomainAllowed,
  recipientsWithDisallowedDomains,
  getAllowedDomains,
};
