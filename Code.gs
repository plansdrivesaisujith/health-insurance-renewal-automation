/**
 * Health Insurance Renewal Automation
 * Google Sheets + Google Apps Script
 *
 * Core workflow:
 * Policies -> daily check -> email template -> Gmail -> Email Log
 *
 * No external API is required for email.
 */

const CONFIG = {
  POLICIES_SHEET: "Policies",
  LOG_SHEET: "Email Log",
  TEMPLATES_SHEET: "Email Templates",
  DEFAULT_REMINDER_SCHEDULE: "60,45,30,21,14,7,3,1",
  DEFAULT_TIME_ZONE: Session.getScriptTimeZone()
};

/**
 * One-time setup.
 * Creates/repairs the required sheets, headers, and daily trigger.
 *
 * Run this once after connecting the Apps Script project to your Google Sheet.
 */
function setupAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensurePoliciesSheet_(ss);
  ensureLogSheet_(ss);
  ensureTemplatesSheet_(ss);
  createDailyTrigger_();

  Logger.log("Setup completed.");
}

/**
 * Creates the daily production trigger.
 * It runs once per day and calls runDailyRenewalCheck().
 */
function createDailyTrigger_() {
  const functionName = "runDailyRenewalCheck";

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log("Daily trigger created for approximately 09:00.");
}

/**
 * Production function.
 *
 * Checks every policy and sends the appropriate reminder once.
 * Duplicate protection is based on:
 * Policy Number + Renewal Date + Reminder Type
 */
function runDailyRenewalCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const policiesSheet = ss.getSheetByName(CONFIG.POLICIES_SHEET);
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  const templatesSheet = ss.getSheetByName(CONFIG.TEMPLATES_SHEET);

  if (!policiesSheet || !logSheet || !templatesSheet) {
    throw new Error(
      "Required sheets are missing. Run setupAutomation() first."
    );
  }

  const policyData = policiesSheet.getDataRange().getValues();

  if (policyData.length < 2) {
    Logger.log("No policy records found.");
    return;
  }

  const headers = policyData[0].map(function(header) {
    return String(header).trim();
  });

  const headerMap = {};
  headers.forEach(function(header, index) {
    headerMap[header] = index;
  });

  validateRequiredHeaders_(headerMap);

  const templates = loadTemplates_(templatesSheet);
  const sentKeys = loadSentKeys_(logSheet);

  let sentCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  for (let rowIndex = 1; rowIndex < policyData.length; rowIndex++) {
    const row = policyData[rowIndex];
    const policy = rowToObject_(headers, row);

    try {
      if (String(policy["Auto Email"]).trim().toUpperCase() !== "YES") {
        skippedCount++;
        continue;
      }

      if (!policy["Email"]) {
        throw new Error("Customer Email is empty.");
      }

      if (!policy["End Date"]) {
        throw new Error("End Date is empty.");
      }

      const expiryDate = normalizeDate_(policy["End Date"]);
      const today = startOfDay_(new Date());

      const daysRemaining = Math.round(
        (expiryDate.getTime() - today.getTime()) / 86400000
      );

      let reminderType = null;

      if (daysRemaining < 0) {
        reminderType = "EXPIRED";
      } else {
        const schedule = parseReminderSchedule_(
          policy["Reminder Schedule"] || CONFIG.DEFAULT_REMINDER_SCHEDULE
        );

        if (schedule.indexOf(daysRemaining) !== -1) {
          reminderType = String(daysRemaining);
        }
      }

      if (!reminderType) {
        skippedCount++;
        continue;
      }

      const policyNumber = String(policy["Policy Number"] || "").trim();
      const renewalKey = formatDate_(expiryDate);
      const duplicateKey =
        policyNumber + "|" + renewalKey + "|" + reminderType;

      if (sentKeys.has(duplicateKey)) {
        skippedCount++;
        continue;
      }

      const template = templates[reminderType];

      if (!template || !template.active) {
        throw new Error(
          "No active email template found for reminder type: " +
          reminderType
        );
      }

      const values = buildReplacementValues_(policy, expiryDate, daysRemaining);

      const subject = replacePlaceholders_(template.subject, values);
      const body = replacePlaceholders_(template.body, values);

      const options = {};

      const advisorEmail = String(policy["Advisor Email"] || "").trim();

      if (advisorEmail) {
        options.cc = advisorEmail;
      }

      GmailApp.sendEmail(
        String(policy["Email"]).trim(),
        subject,
        body,
        options
      );

      appendLog_(
        logSheet,
        policy,
        reminderType,
        daysRemaining,
        "SENT",
        ""
      );

      updatePolicyStatus_(
        policiesSheet,
        rowIndex + 1,
        headerMap,
        reminderType
      );

      sentKeys.add(duplicateKey);
      sentCount++;

    } catch (error) {
      errorCount++;

      appendLog_(
        logSheet,
        policy,
        "ERROR",
        "",
        "ERROR",
        error.message
      );

      Logger.log(
        "Error processing row " + (rowIndex + 1) + ": " + error.message
      );
    }
  }

  Logger.log(
    "Daily renewal check completed. Sent: " +
      sentCount +
      ", Skipped: " +
      skippedCount +
      ", Errors: " +
      errorCount
  );
}

/**
 * Safe manual test for the current first policy.
 *
 * This sends an email only if the first policy is currently due for
 * one of its configured reminder days.
 */
function testFirstPolicyNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const policiesSheet = ss.getSheetByName(CONFIG.POLICIES_SHEET);
  const templatesSheet = ss.getSheetByName(CONFIG.TEMPLATES_SHEET);

  if (!policiesSheet || !templatesSheet) {
    throw new Error("Run setupAutomation() first.");
  }

  const data = policiesSheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error("No customer found in Policies.");
  }

  const headers = data[0].map(function(header) {
    return String(header).trim();
  });

  const policy = rowToObject_(headers, data[1]);

  if (String(policy["Auto Email"]).trim().toUpperCase() !== "YES") {
    throw new Error("Auto Email is not YES for the first customer.");
  }

  const expiryDate = normalizeDate_(policy["End Date"]);
  const today = startOfDay_(new Date());

  const daysRemaining = Math.round(
    (expiryDate.getTime() - today.getTime()) / 86400000
  );

  Logger.log(
    "Customer: " +
      policy["Customer Name"] +
      " | Expiry: " +
      formatDate_(expiryDate) +
      " | Days remaining: " +
      daysRemaining
  );

  if (daysRemaining < 0) {
    Logger.log("Policy is expired. Production function would use EXPIRED.");
    return;
  }

  const schedule = parseReminderSchedule_(
    policy["Reminder Schedule"] || CONFIG.DEFAULT_REMINDER_SCHEDULE
  );

  if (schedule.indexOf(daysRemaining) === -1) {
    Logger.log(
      "No reminder is due today. Current schedule: " +
        schedule.join(",")
    );
    return;
  }

  Logger.log(
    "Reminder due: " + daysRemaining + " Day Reminder."
  );

  // Deliberately calls the production function so duplicate protection
  // and logging are used exactly as they will be in the daily trigger.
  runDailyRenewalCheck();
}

/**
 * Optional helper: generates a WhatsApp-ready link for a policy.
 *
 * This does NOT automatically send WhatsApp messages.
 * It opens a WhatsApp chat with the message pre-filled.
 */
function createWhatsAppLinkForFirstPolicy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.POLICIES_SHEET);

  if (!sheet) throw new Error("Policies sheet not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("No customer found.");

  const headers = data[0].map(function(header) {
    return String(header).trim();
  });

  const policy = rowToObject_(headers, data[1]);

  const phone = String(policy["Customer WhatsApp"] || "").replace(/\D/g, "");

  if (!phone) {
    throw new Error(
      "Customer WhatsApp is empty. Add a Customer WhatsApp column and number."
    );
  }

  const expiryDate = normalizeDate_(policy["End Date"]);
  const today = startOfDay_(new Date());

  const daysRemaining = Math.round(
    (expiryDate.getTime() - today.getTime()) / 86400000
  );

  const message =
    "Hi " + (policy["Customer Name"] || "Customer") + ",\n\n" +
    "Your " + (policy["Insurance Company"] || "") +
    " - " + (policy["Policy Name"] || "") +
    " health insurance policy is approaching its renewal date.\n\n" +
    "Policy Number: " + (policy["Policy Number"] || "") + "\n" +
    "Renewal Date: " + formatDate_(expiryDate) + "\n" +
    "Days Remaining: " + daysRemaining + "\n\n" +
    "If you would like, I can review your existing coverage and help " +
    "you with your renewal options.\n\n" +
    "Regards,\n" +
    (policy["Advisor Name"] || "Your Health Insurance Advisor") + "\n" +
    (policy["Advisor Email"] || "");

  const url =
    "https://wa.me/" +
    phone +
    "?text=" +
    encodeURIComponent(message);

  Logger.log(url);
  return url;
}

/* ------------------------- Helpers ------------------------- */

function ensurePoliciesSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.POLICIES_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.POLICIES_SHEET);
  }

  const headers = [
    "Customer Name",
    "Email",
    "Insurance Company",
    "Policy Number",
    "Policy Name",
    "Add Ons",
    "Start Date",
    "End Date",
    "Advisor Email",
    "Auto Email",
    "Reminder Schedule",
    "Expired Sent",
    "Last Email Sent",
    "Last Email Type",
    "Renewal Status",
    "Notes"
  ];

  const currentHeaders = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length))
    .getValues()[0];

  if (currentHeaders.every(function(value) { return value === ""; })) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function ensureLogSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET);
  }

  const headers = [
    "Date/Time",
    "Customer Name",
    "Customer Email",
    "Insurance Company",
    "Policy Number",
    "Reminder Type",
    "Days Remaining",
    "Status",
    "Error"
  ];

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  if (current.every(function(value) { return value === ""; })) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function ensureTemplatesSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.TEMPLATES_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TEMPLATES_SHEET);
  }

  const headers = ["Template ID", "Subject", "Body", "Active"];

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  if (current.every(function(value) { return value === ""; })) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  if (sheet.getLastRow() === 1) {
    const body30 =
      "Dear {{customer_name}},\n\n" +
      "Your {{insurance_company}} - {{policy_name}} health insurance policy " +
      "is approaching its renewal date, with approximately " +
      "{{days_remaining}} days remaining.\n\n" +
      "Policy Details:\n\n" +
      "Insurance Company: {{insurance_company}}\n" +
      "Policy Name: {{policy_name}}\n" +
      "Policy Number: {{policy_number}}\n" +
      "Renewal Date: {{renewal_date}}\n" +
      "Days Remaining: {{days_remaining}}\n" +
      "Add-ons: {{addons}}\n\n" +
      "Before renewing, I recommend reviewing your current coverage and " +
      "benefits to make sure they still meet your requirements.\n\n" +
      "If you would like, I can review your existing policy and help you " +
      "understand whether there are any suitable coverage improvements, " +
      "upgrades, or alternative options available in the market based on " +
      "your requirements.\n\n" +
      "Please feel free to contact me if you would like to discuss your renewal.\n\n" +
      "Regards,\n" +
      "{{advisor_name}}\n" +
      "Health Insurance Advisor\n" +
      "{{advisor_email}}";

    sheet.getRange(2, 1, 1, 4).setValues([[
      "30",
      "Your {{insurance_company}} Health Insurance Renewal - 30 Days Remaining",
      body30,
      "YES"
    ]]);
  }
}

function validateRequiredHeaders_(headerMap) {
  const required = [
    "Customer Name",
    "Email",
    "Insurance Company",
    "Policy Number",
    "Policy Name",
    "Add Ons",
    "Start Date",
    "End Date",
    "Advisor Email",
    "Auto Email",
    "Reminder Schedule"
  ];

  required.forEach(function(header) {
    if (headerMap[header] === undefined) {
      throw new Error(
        'Missing required Policies column: "' + header + '"'
      );
    }
  });
}

function rowToObject_(headers, row) {
  const object = {};

  headers.forEach(function(header, index) {
    object[header] = row[index];
  });

  return object;
}

function loadTemplates_(sheet) {
  const data = sheet.getDataRange().getValues();
  const templates = {};

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim();

    if (!id) continue;

    templates[id] = {
      subject: String(data[i][1] || ""),
      body: String(data[i][2] || ""),
      active: String(data[i][3] || "")
        .trim()
        .toUpperCase() === "YES"
    };
  }

  return templates;
}

function loadSentKeys_(sheet) {
  const data = sheet.getDataRange().getValues();
  const keys = new Set();

  if (data.length < 2) return keys;

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][7] || "").trim().toUpperCase();

    if (status !== "SENT") continue;

    const policyNumber = String(data[i][4] || "").trim();
    const reminderType = String(data[i][5] || "").trim();

    // Existing logs from older versions may not contain renewal date.
    // New production logs use the duplicate key stored in Notes.
    const notes = String(data[i][8] || "");

    const match = notes.match(/DuplicateKey=(.+)$/);

    if (match) {
      keys.add(match[1]);
    } else if (policyNumber && reminderType) {
      // Conservative fallback for old log rows.
      keys.add(policyNumber + "||" + reminderType);
    }
  }

  return keys;
}

function buildReplacementValues_(policy, expiryDate, daysRemaining) {
  return {
    "{{customer_name}}": policy["Customer Name"] || "",
    "{{insurance_company}}": policy["Insurance Company"] || "",
    "{{policy_name}}": policy["Policy Name"] || "",
    "{{policy_number}}": policy["Policy Number"] || "",
    "{{renewal_date}}": formatDate_(expiryDate),
    "{{days_remaining}}": daysRemaining,
    "{{addons}}": policy["Add Ons"] || "None",
    "{{advisor_name}}":
      policy["Advisor Name"] || "Your Health Insurance Advisor",
    "{{advisor_email}}": policy["Advisor Email"] || ""
  };
}

function replacePlaceholders_(text, values) {
  let result = String(text || "");

  Object.keys(values).forEach(function(placeholder) {
    result = result.split(placeholder).join(String(values[placeholder]));
  });

  return result;
}

function appendLog_(
  sheet,
  policy,
  reminderType,
  daysRemaining,
  status,
  errorMessage
) {
  const expiryDate = policy["End Date"]
    ? normalizeDate_(policy["End Date"])
    : null;

  const renewalKey = expiryDate ? formatDate_(expiryDate) : "";

  const duplicateKey =
    String(policy["Policy Number"] || "").trim() +
    "|" +
    renewalKey +
    "|" +
    reminderType;

  const notes =
    status === "SENT"
      ? "DuplicateKey=" + duplicateKey
      : "";

  sheet.appendRow([
    new Date(),
    policy["Customer Name"] || "",
    policy["Email"] || "",
    policy["Insurance Company"] || "",
    policy["Policy Number"] || "",
    reminderType,
    daysRemaining,
    status,
    errorMessage || notes
  ]);

  // For SENT rows, put the duplicate key in the Error/Notes column.
  // This preserves the user's existing 9-column Email Log layout.
  if (status === "SENT" && notes) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 9).setValue(notes);
  }
}

function updatePolicyStatus_(
  sheet,
  rowNumber,
  headerMap,
  reminderType
) {
  const now = new Date();

  if (headerMap["Last Email Sent"] !== undefined) {
    sheet
      .getRange(rowNumber, headerMap["Last Email Sent"] + 1)
      .setValue(now);
  }

  if (headerMap["Last Email Type"] !== undefined) {
    sheet
      .getRange(rowNumber, headerMap["Last Email Type"] + 1)
      .setValue(reminderType);
  }

  if (headerMap["Renewal Status"] !== undefined) {
    const status = reminderType === "EXPIRED"
      ? "EXPIRED"
      : "RENEWAL REMINDER SENT";

    sheet
      .getRange(rowNumber, headerMap["Renewal Status"] + 1)
      .setValue(status);
  }

  if (
    reminderType === "EXPIRED" &&
    headerMap["Expired Sent"] !== undefined
  ) {
    sheet
      .getRange(rowNumber, headerMap["Expired Sent"] + 1)
      .setValue("YES");
  }
}

function parseReminderSchedule_(value) {
  return String(value)
    .split(",")
    .map(function(item) {
      return parseInt(String(item).trim(), 10);
    })
    .filter(function(number) {
      return !isNaN(number);
    });
}

function normalizeDate_(value) {
  const date = new Date(value);

  if (isNaN(date.getTime())) {
    throw new Error("Invalid date: " + value);
  }

  return startOfDay_(date);
}

function startOfDay_(date) {
  const result = new Date(date);
  result.setHours(0, 0, 0, 0);
  return result;
}

function formatDate_(date) {
  return Utilities.formatDate(
    date,
    CONFIG.DEFAULT_TIME_ZONE,
    "dd-MM-yyyy"
  );
}
