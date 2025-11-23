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
      MailboxEnums: {
        ItemNotificationMessageType: {
          InformationalMessage: "informationalMessage",
        },
      },
      context: {
        mailbox: {
          item: {
            getAttachmentsAsync: jest.fn((cb) =>
              cb({ status: "succeeded", value: [] }),
            ),
            notificationMessages: {},
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
    Office.context.mailbox.item.notificationMessages.replaceAsync = jest.fn(
      (id, options, cb) => cb({ status: "succeeded" }),
    );

    await onMessageSend(event);

    const notificationType =
      Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;
    expect(
      Office.context.mailbox.item.notificationMessages.replaceAsync,
    ).toHaveBeenCalledWith(
      "AttachmentBlock",
      expect.objectContaining({
        message: expect.any(String),
        icon: "icon32",
        persistent: true,
        type: notificationType,
      }),
      expect.any(Function),
    );
    expect(event.completed).toHaveBeenCalledWith({ allowEvent: false });
  });
});
