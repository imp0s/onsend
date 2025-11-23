const { onMessageSend, getAttachments } = require("../src/addin");

describe("getAttachments", () => {
  beforeEach(() => {
    global.Office = {
      AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      MailboxEnums: {},
      context: {
        mailbox: {
          item: {
            attachments: [{ id: "a1" }],
          },
        },
      },
    };
  });

  afterEach(() => {
    delete global.Office;
  });

  it("returns cached attachments when async API is unavailable", async () => {
    const result = await getAttachments();
    expect(result).toEqual([{ id: "a1" }]);
  });

  it("uses getAttachmentsAsync when available", async () => {
    const attachments = [{ id: "a2" }];
    Office.context.mailbox.item.getAttachmentsAsync = jest.fn((cb) =>
      cb({ status: "succeeded", value: attachments }),
    );

    const result = await getAttachments();
    expect(result).toEqual(attachments);
  });
});

describe("onMessageSend", () => {
  let event;

  beforeEach(() => {
    event = { completed: jest.fn() };
    global.Office = {
      AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      MailboxEnums: {},
      context: {
        mailbox: {
          item: {
            getAttachmentsAsync: jest.fn((cb) =>
              cb({ status: "succeeded", value: [] }),
            ),
            notificationMessages: {
              addAsync: jest.fn((id, options, cb) =>
                cb({ status: "succeeded" }),
              ),
            },
          },
        },
      },
    };
  });

  afterEach(() => {
    delete global.Office;
  });

  it("allows send when there are no attachments", async () => {
    await onMessageSend(event);
    expect(event.completed).toHaveBeenCalledWith({ allowEvent: true });
  });

  it("blocks send and shows notification when attachments exist", async () => {
    Office.context.mailbox.item.getAttachmentsAsync = jest.fn((cb) =>
      cb({ status: "succeeded", value: [{ id: "a1" }] }),
    );

    await onMessageSend(event);

    expect(event.completed).toHaveBeenCalledWith({
      allowEvent: false,
      errorMessage: expect.any(String),
    });
  });
});
