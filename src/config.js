// Domains allowed in addition to the sender's domain when validating recipients.
// Use bare domain names (e.g., "example.com") or suffixes (e.g., "example.co.uk").
const allowedDomainExtensions = ["example.com"];

// Allowed attachment file extensions (case-insensitive).
// If empty, all attachments are blocked.
const allowedAttachmentExtensions = [".png", ".jpg", ".jpeg", ".gif", ".txt"];

module.exports = { allowedDomainExtensions, allowedAttachmentExtensions };
