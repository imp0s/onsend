const {
  onMessageSend,
  getAttachments,
  showBlockingNotification,
} = require("../src/addin");

describe("getAttachments", () => {
  beforeEach(() => {
    global.Office = {
      AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      MailboxEnums: {
        ItemNotificationMessageType: { InformationalMessage: "inform" },
      },
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
      MailboxEnums: {
        ItemNotificationMessageType: { InformationalMessage: "inform" },
      },
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

    expect(
      Office.context.mailbox.item.notificationMessages.addAsync,
    ).toHaveBeenCalledWith(
      "AttachmentBlock",
      expect.objectContaining({ message: expect.any(String) }),
      expect.any(Function),
    );
    expect(event.completed).toHaveBeenCalledWith({ allowEvent: false });
  });
});

describe("showBlockingNotification", () => {
  beforeEach(() => {
    global.Office = {
      AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      MailboxEnums: {
        ItemNotificationMessageType: { InformationalMessage: "inform" },
      },
      context: {
        mailbox: {
          item: {
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

  it("resolves when notification is added", async () => {
    await expect(showBlockingNotification()).resolves.toBeUndefined();
  });

  it("rejects when notification fails", async () => {
    Office.context.mailbox.item.notificationMessages.addAsync = jest.fn(
      (id, opts, cb) => cb({ status: "failed", error: new Error("fail") }),
    );

    await expect(showBlockingNotification()).rejects.toThrow("fail");
  });
});
