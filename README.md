# Health Insurance Renewal Automation

**Implemented by:** Saisujith Vanimisetti  
**Project:** Health Insurance Renewal Automation  
**Platform:** Google Sheets + Google Apps Script  
**Repository:** GitHub-ready reusable automation


A reusable Google Sheets + Google Apps Script system for automated health-insurance renewal reminders.

## What it does

- Reads policy/customer information from a Google Sheet.
- Checks policies once per day.
- Supports configurable reminder days such as:
  `60,45,30,21,14,7,3,1`
- Sends a separate email to each customer.
- CCs the advisor on each email.
- Uses editable email templates stored in the `Email Templates` sheet.
- Logs every sent email in `Email Log`.
- Prevents duplicate reminders.
- Sends one expired notification when a policy becomes expired.
- Does not require Zapier, Make, or another external automation service.
- Includes an optional WhatsApp-ready link helper, but does not automatically send WhatsApp messages.

## Repository structure

```text
health-insurance-renewal-automation/
├── Code.gs
├── appsscript.json
├── README.md
└── sample/
    ├── Policies.csv
    ├── Email Log.csv
    └── Email Templates.csv
```

## Google Sheet structure

### Policies

Use these columns:

| Column | Header |
|---|---|
| A | Customer Name |
| B | Email |
| C | Insurance Company |
| D | Policy Number |
| E | Policy Name |
| F | Add Ons |
| G | Start Date |
| H | End Date |
| I | Advisor Email |
| J | Auto Email |
| K | Reminder Schedule |
| L | Expired Sent |
| M | Last Email Sent |
| N | Last Email Type |
| O | Renewal Status |
| P | Notes |

Example:

```text
Customer Name: Test Customer
Email: your-email@example.com
Insurance Company: Care Health Insurance
Policy Number: C4548020
Policy Name: Care Supreme
Add Ons: None
Start Date: 04-10-2026
End Date: 03-11-2026
Advisor Email: your-email@example.com
Auto Email: YES
Reminder Schedule: 60,45,30,21,14,7,3,1
```

### Email Log

```text
Date/Time
Customer Name
Customer Email
Insurance Company
Policy Number
Reminder Type
Days Remaining
Status
Error
```

### Email Templates

```text
Template ID
Subject
Body
Active
```

Template IDs are normally the number of days remaining:

```text
60
45
30
21
14
7
3
1
```

Use `EXPIRED` for the expired-policy template.

## Supported placeholders

Email subjects and bodies can use:

```text
{{customer_name}}
{{insurance_company}}
{{policy_name}}
{{policy_number}}
{{renewal_date}}
{{days_remaining}}
{{addons}}
{{advisor_name}}
{{advisor_email}}
```

If `Advisor Name` is not yet a column in Policies, the system uses:

```text
Your Health Insurance Advisor
```

## Initial setup

### Option 1 — easiest: Apps Script editor

1. Create/open your Google Sheet.
2. Open **Extensions → Apps Script**.
3. Replace the code with `Code.gs`.
4. Open **Project Settings** and set the time zone to:
   `Europe/Amsterdam`
5. Save.
6. Run:

```text
setupAutomation
```

7. Authorize the script when Google asks.
8. Run:

```text
testFirstPolicyNow
```

The test customer should only receive an email if the current day matches their configured reminder schedule.

### Daily automation

`setupAutomation()` creates a time-driven trigger for:

```text
runDailyRenewalCheck()
```

The trigger runs once per day at approximately 09:00 in the project time zone.

You do not need your laptop to be switched on.

## Duplicate protection

A sent reminder is identified using:

```text
Policy Number + Renewal Date + Reminder Type
```

Example:

```text
C4548020|03-11-2026|30
```

Therefore two customers with the same renewal date still receive separate emails.

The system never clubs customers together and does not BCC multiple customers.

## Important test safety

Do not create a trigger for:

```text
test30DayTemplate()
```

That function was only for testing the template.

The production trigger must call:

```text
runDailyRenewalCheck()
```

## WhatsApp

The repository includes:

```text
createWhatsAppLinkForFirstPolicy()
```

This creates a WhatsApp `wa.me` link with a personalized message.

It does **not** automatically send the message.

The normal personal WhatsApp desktop app does not provide a supported interface for Google Apps Script to automatically press Send. Automatic WhatsApp sending requires an official WhatsApp API or another supported integration.

## SMS

SMS is not provided by Google Apps Script itself. Automatic SMS delivery normally requires an SMS provider and therefore is not included in the free core system.

## Product architecture

```text
                    Google Sheet
                         |
                         v
                runDailyRenewalCheck
                         |
              +----------+----------+
              |                     |
         Renewal due?           Not due
              |                     |
              v                     v
        Load template              Stop
              |
              v
       Replace placeholders
              |
              v
          Gmail Send
              |
              v
          Email Log
              |
              v
       Duplicate protection
```

## Future extensions

The architecture can later support:

- Advisor Name column
- Settings sheet
- More email templates
- HTML email templates
- WhatsApp API
- SMS provider
- Dashboard
- Multiple advisors
- Policy categories
- Customer status
- Renewal conversion tracking
- Deployment/versioning through Git + clasp

## Security

Do not commit the following to Git:

- API tokens
- Passwords
- Customer data
- Real customer email addresses
- Real policy numbers
- OAuth credentials
- Service-account keys

Use sample data in the repository.
