describe("domain safety helpers", () => {
  let addin;
  let primaryDomain;

  beforeEach(() => {
    ({ allowedDomainExtensions: [primaryDomain] = [] } = require("../src/config"));
    if (!primaryDomain) {
      throw new Error("Configure at least one allowed domain in src/config.js");
    }

    global.Office = {
      context: {
        mailbox: {
          userProfile: { emailAddress: `sender@${primaryDomain}` },
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
    expect(addin.extractDomain(`user@${primaryDomain.toUpperCase()}`)).toBe(
      primaryDomain,
    );
  });

  test("getAllowedDomains returns sender domain when available", () => {
    const { domains, enforceExact } = addin.getAllowedDomains();
    expect(domains).toEqual([primaryDomain]);
    expect(enforceExact).toBe(true);
  });

  test("isDomainAllowed enforces exact match when sender domain is known", () => {
    const { domains, enforceExact } = addin.getAllowedDomains();
    expect(addin.isDomainAllowed(primaryDomain, domains, enforceExact)).toBe(true);
    expect(
      addin.isDomainAllowed("other-example.com", domains, enforceExact),
    ).toBe(false);
  });

  test("recipientsWithDisallowedDomains flags mismatches", () => {
    const { domains, enforceExact } = addin.getAllowedDomains();
    const recipients = [
      { emailAddress: `teammate@${primaryDomain}` },
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
