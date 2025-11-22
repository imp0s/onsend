const { normalizeAttachmentId } = require("../src/addin");

describe("normalizeAttachmentId", () => {
  it("prefers id over attachmentId when both are present", () => {
    const attachment = {
      id: "primary-id",
      attachmentId: "fallback-id",
    };
    expect(normalizeAttachmentId(attachment)).toBe("primary-id");
  });

  it("falls back to attachmentId when id is missing", () => {
    const attachment = {
      attachmentId: "fallback-only",
    };
    expect(normalizeAttachmentId(attachment)).toBe("fallback-only");
  });
});
