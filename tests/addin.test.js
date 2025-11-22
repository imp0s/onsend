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
    global.Office = {
      AsyncResultStatus: { Succeeded: "succeeded" },
      MailboxEnums: {
        AttachmentContentFormat: { Base64: "base64", FileUrl: "fileUrl" },
      },
      context: {
        mailbox: {
          item: {
            getAttachmentContentAsync: jest.fn(),
          },
          makeEwsRequestAsync: jest.fn((request, cb) => {
            const ewsResponse = `<?xml version="1.0" encoding="utf-8"?>\n<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">\n  <soap:Body>\n    <m:GetAttachmentResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">\n      <m:ResponseMessages>\n        <m:GetAttachmentResponseMessage ResponseClass="Success">\n          <m:Attachments>\n            <t:FileAttachment><t:Content>${base64Dummy}</t:Content></t:FileAttachment>\n          </m:Attachments>\n        </m:GetAttachmentResponseMessage>\n      </m:ResponseMessages>\n    </m:GetAttachmentResponse>\n  </soap:Body>\n</soap:Envelope>`;

            cb({
              status: "succeeded",
              value: ewsResponse,
            });
          }),
        },
      },
    };
  });

  afterEach(() => {
    jest.resetAllMocks();
    delete global.Office;
  });

  it("downloads fileUrl content via EWS and returns base64", async () => {
    const result = await getAttachmentContent("att-1");

    expect(Office.context.mailbox.makeEwsRequestAsync).toHaveBeenCalledWith(
      expect.stringContaining("GetAttachment"),
      expect.any(Function),
    );

    expect(
      Office.context.mailbox.item.getAttachmentContentAsync,
    ).not.toHaveBeenCalled();

    expect(result).toEqual({
      format: "base64",
      content: base64Dummy,
    });
  });
});
