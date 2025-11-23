/* global Office */

describe("domain safety helpers", () => {
  let addin;

  beforeEach(() => {
    global.Office = {
      context: {
        mailbox: {
          userProfile: { emailAddress: "sender@example.com" },
        },
      },
      MailboxEnums: {
        ItemNotificationMessageType: { InformationalMessage: "info" },
      },
      actions: { associate: jest.fn() },
      onReady: jest.fn((cb) => cb()),
    };

    addin = require("../src/addin");
  });

  afterEach(() => {
    jest.resetModules();
    delete global.Office;
  });

  test("extractDomain returns domain for valid email", () => {
    expect(addin.extractDomain("user@EXAMPLE.com")).toBe("example.com");
  });

  test("getAllowedDomains returns sender domain when available", () => {
    const { domains, enforceExact } = addin.getAllowedDomains();
    expect(domains).toEqual(["example.com"]);
    expect(enforceExact).toBe(true);
  });

  test("isDomainAllowed enforces exact match when sender domain is known", () => {
    const { domains, enforceExact } = addin.getAllowedDomains();
    expect(addin.isDomainAllowed("example.com", domains, enforceExact)).toBe(
      true,
    );
    expect(
      addin.isDomainAllowed("other-example.com", domains, enforceExact),
    ).toBe(false);
  });

  test("recipientsWithDisallowedDomains flags mismatches", () => {
    const { domains, enforceExact } = addin.getAllowedDomains();
    const recipients = [
      { emailAddress: "teammate@example.com" },
      { emailAddress: "external@outside.com" },
    ];

    const offenders = addin.recipientsWithDisallowedDomains(
      recipients,
      domains,
      enforceExact,
    );

    expect(offenders).toHaveLength(1);
    expect(offenders[0].domain).toBe("outside.com");
  });
});
