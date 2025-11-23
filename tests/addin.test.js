describe("domain safety helpers with configured allowances", () => {
  let addin;
  let primaryDomain;
  const allowedDomainExtensions = ["example.com", "trusted.org"];

  beforeEach(() => {
    jest.resetModules();
    global.EmailSafetyConfig = { allowedDomainExtensions };

    ({
      allowedDomainExtensions: [primaryDomain] = [],
    } = global.EmailSafetyConfig);

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

    addin = require("../public/app");
  });

  afterEach(() => {
    delete global.Office;
    delete global.EmailSafetyChecker;
    delete global.EmailSafetyConfig;
  });

  test("extractDomain returns domain for valid email", () => {
    expect(addin.extractDomain(`user@${primaryDomain.toUpperCase()}`)).toBe(
      primaryDomain,
    );
  });

  test("getAllowedDomains returns sender and configured domains", () => {
    const { domains, enforceExact } = addin.getAllowedDomains();
    expect(domains).toEqual([
      ...new Set([primaryDomain, ...allowedDomainExtensions]),
    ]);
    expect(enforceExact).toBe(false);
  });

  test("isDomainAllowed allows configured domains alongside sender domain", () => {
    const { domains, enforceExact } = addin.getAllowedDomains();
    expect(addin.isDomainAllowed(primaryDomain, domains, enforceExact)).toBe(
      true,
    );
    expect(addin.isDomainAllowed("trusted.org", domains, enforceExact)).toBe(
      true,
    );
    expect(addin.isDomainAllowed("unlisted.net", domains, enforceExact)).toBe(
      false,
    );
  });

  test("recipientsWithDisallowedDomains flags mismatches", () => {
    const { domains, enforceExact } = addin.getAllowedDomains();
    const recipients = [
      { emailAddress: `teammate@${primaryDomain}` },
      { emailAddress: "external@unlisted.net" },
    ];

    const offenders = addin.recipientsWithDisallowedDomains(
      recipients,
      domains,
      enforceExact,
    );

    expect(offenders).toHaveLength(1);
    expect(offenders[0].domain).toBe("unlisted.net");
  });
});

describe("domain safety helpers without configured allowances", () => {
  let addin;
  const senderDomain = "solo.com";

  beforeEach(() => {
    jest.resetModules();
    global.EmailSafetyConfig = { allowedDomainExtensions: [] };

    global.Office = {
      context: {
        mailbox: {
          userProfile: { emailAddress: `sender@${senderDomain}` },
        },
      },
      MailboxEnums: {
        ItemNotificationMessageType: { InformationalMessage: "info" },
      },
      actions: { associate: jest.fn() },
      onReady: jest.fn((cb) => cb()),
    };

    addin = require("../public/app");
  });

  afterEach(() => {
    delete global.Office;
    delete global.EmailSafetyChecker;
    delete global.EmailSafetyConfig;
  });

  test("getAllowedDomains enforces exact sender match when no config provided", () => {
    const { domains, enforceExact } = addin.getAllowedDomains();
    expect(domains).toEqual([senderDomain]);
    expect(enforceExact).toBe(true);
    expect(
      addin.isDomainAllowed(`sub.${senderDomain}`, domains, enforceExact),
    ).toBe(false);
  });
});
