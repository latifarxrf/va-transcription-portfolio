# Standard Operating Procedure (SOP): Executive Email Triage & Inbox Zero System

**Author:** Latifar (@latifarxrf)  
**Role:** Executive Virtual Assistant  
**Target:** Senior Executives receiving 50–100 emails/day  
**Goal:** Maintain Inbox Zero daily, 0 missed client inquiries, response latency under 2 hours during business hours.

---

## 1. Core Triage Workflow

```
Incoming Email
      │
      ├─► Is it Spam/Cold Pitch/Newsletter? ──► [Auto-Filter] ──► Archive / Unsubscribe
      │
      ├─► Is it Urgent (Client, Revenue, Escaped Issue)?
      │     └─► Label: 🔴 01-Urgent-Client
      │     └─► Notify via Slack / WhatsApp (if outside deep work)
      │     └─► Action SLA: < 1 Hour
      │
      ├─► Does it require Executive input or decision?
      │     └─► Label: 🟡 02-Action-Required
      │     └─► Draft reply in Canned Templates
      │     └─► Action SLA: End of Day
      │
      ├─► Is it an informational update / receipt / report?
      │     └─► Label: 🟢 03-FYI-Read-Later
      │     └─► Move to Weekly Digest
      │
      └─► Is it meeting scheduling?
            └─► Cross-reference Calendar Blocking Rules
            └─► Reply with Calendly link or 3 proposed slots
```

---

## 2. Labeling Architecture & Color Codes

| Label Name | Color | Trigger Criteria | SLA |
|---|---|---|---|
| `🔴 01-Urgent-Client` | Red | Active paying client, contract renewal, payment issue | < 1 hour |
| `🟡 02-Action-Required` | Yellow | Vendor questions, team approvals, direct requests | < 4 hours |
| `🔵 03-Waiting-Response` | Blue | Sent emails awaiting third-party reply | Review 48h |
| `🟢 04-FYI-Read-Later` | Green | Weekly newsletters, industry reports, company-wide updates | Weekly batch |
| `📁 05-Receipts-Invoices` | Grey | Billing receipts, subscriptions, statements | Auto-archive to Drive |

---

## 3. Canned Response Templates

### Template A: Meeting Request (Availability Offered)
```text
Subject: Re: Meeting Request — [Topic] / [Client Name]

Hi [Name],

Thanks for reaching out. [Executive Name] would be glad to connect regarding [Topic].

Here are a few 30-minute windows that work with their schedule next week (all times in [Client Timezone]):
1. Tuesday, [Date] at 10:00 AM – 10:30 AM [Timezone]
2. Wednesday, [Date] at 2:00 PM – 2:30 PM [Timezone]
3. Thursday, [Date] at 11:30 AM – 12:00 PM [Timezone]

Alternatively, feel free to pick any slot that best fits your calendar directly here: [Calendly Link].

Looking forward to confirming,

Best regards,
Latifar
Executive Assistant to [Executive Name]
```

### Template B: Gentle Follow-Up (No Reply in 48h)
```text
Subject: Follow-up: [Original Subject]

Hi [Name],

I wanted to quickly follow up on the note below regarding [brief description/deliverable]. 

Please let us know if you need any additional information from our end to proceed, or if there's an updated timeline we should keep in mind.

Thank you,
Latifar
```

### Template C: Delegated Task Handover
```text
Subject: Intro: [Topic] — Connecting you with [Team Member]

Hi [Name],

Thanks for the update. To ensure this is handled promptly, I've looped in [Team Member], who leads our [Department/Function].

[Team Member Name] will take over from here and provide the details you need.

Best regards,
Latifar
```

---

## 4. End-of-Day Checklist
- [ ] Primary inbox count = 0
- [ ] All `🔴 01-Urgent` threads resolved or escalated
- [ ] Draft responses queued in `🟡 02-Action-Required` for executive signature
- [ ] `🔵 03-Waiting` threads older than 4 days followed up
- [ ] Daily time spent on triage logged in tracking tool
