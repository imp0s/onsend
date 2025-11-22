const { normalizeAttachmentId, getAttachmentContent } = require("../src/addin");

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

describe("getAttachmentContent", () => {
  const base64Dummy = Buffer.from("dummy").toString("base64");

  beforeEach(() => {
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      arrayBuffer: async () => new TextEncoder().encode("dummy").buffer,
    });

    global.Office = {
      AsyncResultStatus: { Succeeded: "succeeded" },
      MailboxEnums: {
        AttachmentContentFormat: { Base64: "base64", FileUrl: "fileUrl" },
      },
      context: {
        mailbox: {
          getCallbackTokenAsync: jest.fn((opts, cb) =>
            cb({ status: "succeeded", value: "token123" }),
          ),
          item: {
            getAttachmentContentAsync: jest.fn((id, cb) =>
              cb({
                status: "succeeded",
                value: {
                  format: "fileUrl",
                  content: "https://attachment.example.com/file",
                },
              }),
            ),
          },
        },
      },
    };
  });

  afterEach(() => {
    jest.resetAllMocks();
    delete global.Office;
    delete global.fetch;
  });

  it("downloads fileUrl content with the callback token and returns base64", async () => {
    const result = await getAttachmentContent("att-1");

    expect(global.fetch).toHaveBeenCalledWith(
      "https://attachment.example.com/file",
      expect.objectContaining({
        headers: { Authorization: "Bearer token123" },
        mode: "cors",
        credentials: "include",
      }),
    );

    expect(result).toEqual({
      format: "base64",
      content: base64Dummy,
    });
  });
});
