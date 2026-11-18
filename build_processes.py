#!/usr/bin/env python3
"""Builds the 'Processes' pages: step-by-step guides for the common company-secretarial jobs.

The steps, forms and sections are written by hand in PROCESSES below (edit them there, or add your own in
guide/processes.json, same shape). Drafts from the Drafts library are matched by words in their titles.
Output: site/data/processes.json
"""
import json
import re
import sys
from pathlib import Path

SITE = Path("site")

# s = sections of the Act, f = forms. Time limits are the commonly cited ones; the page tells readers to verify.
PROCESSES = [
    {"id": "incorporation", "title": "Incorporating a company (private limited)",
     "summary": "From name approval to the certificate of incorporation, then the first Board meeting and the filings that follow in the first months.",
     "when": "Starting a new private company through the SPICe+ route.",
     "steps": [
         {"t": "Get digital signatures and DINs", "d": "Every proposed director and subscriber needs a Digital Signature Certificate. DIN is applied for inside the incorporation form for new directors.", "f": ["DIR-3"], "s": ["153", "154"]},
         {"t": "Choose and reserve the name", "d": "Check the name against the Registrar's name rules and trade marks. The name is applied for in Part A of the incorporation form.", "f": ["INC-32"], "s": ["4"]},
         {"t": "Settle the memorandum and articles", "d": "Prepare the Memorandum (objects, liability, capital) and Articles. Use the ready-made forms for the standard version or file your own.", "f": ["INC-33", "INC-34"], "s": ["3", "4", "5"]},
         {"t": "File the incorporation form", "d": "File SPICe+ with its attachments (proof of registered office, director and subscriber documents, declarations) and pay stamp duty and fees.", "f": ["INC-32", "INC-35"], "s": ["7"]},
         {"t": "Receive the certificate of incorporation", "d": "The Registrar issues the certificate with the CIN. PAN and TAN come with it. The date on it is the date of incorporation.", "f": [], "s": ["7", "9"]},
         {"t": "Hold the first Board meeting", "d": "The model agenda has two parts. Administrative: leave of absence, quorum, a chairperson for the meeting. Business: take note of the certificate of incorporation, the Memorandum and Articles and the registered office; confirm the first directors; record directors' disclosures of interest; appoint the Chairperson of the Board; open bank accounts; approve the format of share certificates; apply for PAN, TAN and other registrations; approve preliminary expenses and ratify preliminary contracts; fix the financial year; appoint Key Managerial Personnel; appoint the first auditor; adopt the common seal if one is used.", "f": ["MBP-1"], "s": ["12", "173", "184", "203", "139"], "r": ["A-1.1.*", "A-1.2.*"]},
         {"t": "Display the name and address", "d": "Affix a name board at the registered office and at every other place of business, and print the name, CIN and registered office address on business letters, bills and other papers.", "f": [], "s": ["12"], "r": ["A-1.2.3"]},
         {"t": "Confirm the registered office", "d": "File the notice of the registered office address.", "f": ["INC-22"], "s": ["12"]},
         {"t": "Issue share certificates to subscribers", "d": "Deliver share certificates in the approved format (Form SH-1) to the subscribers within the period the Act gives after incorporation.", "f": ["SH-1"], "s": ["56"], "r": ["A-1.2.8"]},
         {"t": "File the auditor appointment", "d": "The first auditor holds office until the first AGM. Inform the Registrar.", "f": ["ADT-1"], "s": ["139"], "r": ["A-1.2.13"]},
         {"t": "File the declaration for commencement of business", "d": "Directors file a declaration that every subscriber has paid for the shares agreed to be taken.", "f": ["INC-20A"], "s": ["10A"]}],
     "check": ["Digital signatures of all directors and subscribers", "Name available and not identical or too like another", "Draft Memorandum and Articles approved by subscribers", "Proof of registered office address and owner's no-objection", "Identity and address proofs of directors and subscribers", "Stamp duty and fees paid", "Certificate of incorporation received", "First Board meeting held and minuted", "Disclosures of interest (MBP-1) taken from every director", "Name board put up and name, CIN and address on stationery", "First financial year fixed", "Auditor appointed and ADT-1 filed", "Bank account opened", "PAN, TAN and other registrations applied for", "Share certificates printed and issued", "INC-20A filed", "Statutory registers opened"],
     "rs": [{"t": "Agenda of the first Board meeting", "ids": ["A-1.1.*", "A-1.2.*"]}],
     "drafts": ["incorporat", "first board", "subscriber"],
     "lodr": []},
    {"id": "registered-office", "title": "Shifting the registered office",
     "summary": "The route depends on where the office moves: within the same city, elsewhere in the State, to another Registrar's area in the same State, or to another State.",
     "when": "The company wants to change the address or location of its registered office.",
     "steps": [
         {"t": "Decide which case applies", "d": "Four cases: (1) within the local limits of the same city or town; (2) within the State but outside those limits; (3) within the State but to the area of another Registrar; (4) from one State to another. The approvals grow from case to case.", "f": [], "s": ["12", "13"]},
         {"t": "Case 1: within the same city or town", "d": "A Board resolution is enough. It authorises filing of the e-forms, the change of name boards and stationery, and informing banks, customers and the other authorities.", "f": ["INC-22"], "s": ["12"], "r": ["A-3.9.8"]},
         {"t": "Case 2: elsewhere in the State, outside the city limits", "d": "The Board approves, then the members pass a special resolution at a general meeting called under sections 101 and 102. File the notice of the new address.", "f": ["MGT-14", "INC-22"], "s": ["12", "101", "102"], "r": ["A-3.9.7"]},
         {"t": "Case 3: to another Registrar's area in the same State", "d": "Board resolution, special resolution of the members, and confirmation of the Regional Director before the shift. Then the e-forms are filed.", "f": ["MGT-14", "INC-23"], "s": ["12"], "r": ["A-3.9.6"]},
         {"t": "Case 4: from one State to another", "d": "Board resolution; special resolution at a general meeting; alteration of the registered office clause of the Memorandum; notice of the proposed shift published in one English and one vernacular newspaper in the district; application to the Regional Director for confirmation.", "f": ["MGT-14", "INC-23"], "s": ["12", "13", "110"], "r": ["A-3.9.13", "C-2.2.4", "A-3.9.3"]},
         {"t": "Obtain the order and file it", "d": "File the Regional Director's order with the Registrar once received, and give effect to it.", "f": ["INC-28"], "s": ["13"], "r": ["A-3.9.1"]},
         {"t": "Change name boards and stationery", "d": "Change the registered office on the name board, letter heads and official publications, and inform banks, customers, vendors and every authority under whose law the company is registered.", "f": [], "s": ["12"]},
         {"t": "Registers and records elsewhere", "d": "If the company keeps its registers away from the registered office, a resolution is needed for that.", "f": [], "s": ["94"], "r": ["A-3.9.9"]}],
     "check": ["Case worked out: same city, same State, other Registrar, or other State", "Board resolution passed", "General meeting called and special resolution passed where the case needs it", "MGT-14 filed", "Newspaper notices published (shift to another State)", "Application to the Regional Director filed", "Order received and filed", "INC-22 filed", "Name boards, stationery and website changed", "Banks, GST, PAN, licences and customers informed"],
     "rs": [{"t": "Resolutions for shifting the office", "ids": ["A-3.9.8", "A-3.9.7", "A-3.9.6", "A-3.9.13", "A-3.9.1", "A-3.9.3", "A-3.9.9", "C-2.2.4"]}],
     "drafts": ["registered office", "shifting", "change of address"],
     "lodr": [{"id": "30", "why": "Listed companies disclose this event to the stock exchanges."}]},
    {"id": "agm", "title": "Annual General Meeting",
     "summary": "From the Board meeting that approves the accounts to the returns filed after the meeting.",
     "when": "Every year. The Act sets the time within which the meeting must be held.",
     "steps": [
         {"t": "Board meeting that approves the accounts", "d": "Review the compliance status; authorise directors to file e-forms; approve the closure of the register of members; approve the draft Board's report; appoint the secretarial auditor; recommend the dividend; approve the financial statements; recommend the statutory auditor; appoint the cost auditor; approve the draft notice of the AGM.", "f": [], "s": ["96", "134", "91", "123", "139", "148", "173"], "r": ["A-2.1.*"]},
         {"t": "Appoint the scrutiniser and fix the record date", "d": "If e-voting applies, appoint a scrutiniser and fix the cut-off date for voting.", "f": [], "s": ["108"]},
         {"t": "Send the notice", "d": "Send the notice with the annual report to members, directors and the auditor, within the notice period the Act requires. Short notice needs the consent required by the Act. Items of special business carry an explanatory statement.", "f": [], "s": ["101", "102", "136"]},
         {"t": "Hold the meeting", "d": "Check quorum, take the chair, and deal with the usual business: appointment of auditors, declaration of dividend, the financial statements with the reports of the Board and the auditors, appointment of a director in place of one retiring, and ratification of the auditors and their remuneration.", "f": [], "s": ["103", "105", "107", "139", "152"], "r": ["C-1.1.*"]},
         {"t": "Special business", "d": "Items such as capital changes, alteration of the Articles, borrowing limits, related party transactions and appointments of directors go as special business, each with an explanatory statement and the type of resolution it needs.", "f": [], "s": ["102"], "r": ["C-2.*"]},
         {"t": "Prepare the minutes and the scrutiniser's report", "d": "Record the proceedings and the results in the minutes book.", "f": [], "s": ["118"]},
         {"t": "File the special resolutions", "d": "File any resolution that must be filed with the Registrar.", "f": ["MGT-14"], "s": ["117"]},
         {"t": "File the financial statements and the annual return", "d": "File the financial statements and the annual return within the time the Act gives after the meeting.", "f": ["AOC-4", "MGT-7"], "s": ["137", "92"]}],
     "check": ["Financial statements signed by the board and auditor", "Board's report approved", "Register of members closure approved", "Notice and annual report sent in time", "Explanatory statement for each special business item", "Attendance register and proxies kept", "Quorum recorded", "E-voting done where it applies", "Minutes signed", "AOC-4 filed", "MGT-7 filed", "Auditor appointment or reappointment filed, if due"],
     "rs": [{"t": "Board meeting that approves the accounts", "ids": ["A-2.1.*"]}, {"t": "Usual business at the AGM", "ids": ["C-1.1.*"]}],
     "more": [["Special business resolutions for general meetings", "C"]],
     "drafts": ["agm", "annual general", "notice of meeting", "board report", "directors report", "proxy", "attendance"],
     "lodr": [{"id": "29", "why": "Prior intimation of board meetings that consider the accounts."}, {"id": "30", "why": "Disclosure of events and the outcome of meetings."}, {"id": "44", "why": "Voting results of the meeting."}]},
    {"id": "board-meeting", "title": "Board meeting",
     "summary": "Notice, agenda, quorum, minutes and the filings that follow from the decisions.",
     "when": "Each time the board meets. The Act requires a minimum number of meetings a year with a maximum gap between two.",
     "steps": [
         {"t": "Fix the date and prepare the agenda", "d": "Prepare the agenda with the notes for each item. Check the previous minutes for matters that carry over. The model agenda starts with administrative items: leave of absence, confirming the quorum, electing a chairperson for the meeting, and ends with any other item with the permission of the chair.", "f": [], "s": ["173"], "r": ["A-1.1.1", "A-1.1.2", "A-1.1.3", "A-1.1.4"]},
         {"t": "Send the notice", "d": "Send written notice, with agenda and notes, to every director within the notice period the Act requires. Notice may go by electronic means. Hold the meeting at shorter notice only as the Act allows.", "f": [], "s": ["173"], "r": ["A-3.5.2"]},
         {"t": "Check quorum and disclosures", "d": "Confirm the quorum. Take any disclosure of interest at the first meeting of the year and keep interested directors out of the vote.", "f": ["MBP-1"], "s": ["174", "184"], "r": ["A-1.2.5"]},
         {"t": "Conduct the meeting", "d": "Take the chair, pass the resolutions item by item, and record dissent where a director asks. A director may take part by video conference or other audio-visual means.", "f": [], "s": ["175", "173"], "r": ["A-3.5.1"]},
         {"t": "Write and sign the minutes", "d": "Circulate the draft, collect comments, and have the chairman sign within the time the Standards allow. Paste them in the minutes book.", "f": [], "s": ["118"]},
         {"t": "Act on the decisions", "d": "Make the filings that follow, for example resolutions under section 179(3), director changes or allotments.", "f": ["MGT-14", "DIR-12", "PAS-3"], "s": ["117", "179"]}],
     "check": ["Notice and agenda sent in time", "Quorum met", "Disclosures of interest taken", "Resolutions passed and recorded", "Minutes circulated and signed", "Follow-up filings made in time", "Attendance register signed"],
     "rs": [{"t": "Administrative items", "ids": ["A-1.1.*"]}, {"t": "Electronic meetings and notices", "ids": ["A-3.5.*"]}, {"t": "Business items", "ids": ["A-3.7.*"]}],
     "more": [["All Board meeting resolutions", "A"]],
     "drafts": ["board meeting", "board resolution", "minutes", "agenda", "notice of board"],
     "lodr": [{"id": "29", "why": "Prior intimation to the exchange for certain agenda items."}, {"id": "30", "why": "Disclose the outcome within the time the Regulations give."}]},
    {"id": "committees", "title": "Audit Committee and Nomination and Remuneration Committee meetings",
     "summary": "The items each committee takes up, with a model resolution for each.",
     "when": "Each meeting of these committees.",
     "steps": [
         {"t": "Audit Committee: appointments and reviews", "d": "Recommend the appointment of statutory auditors; review internal audit plan, internal controls and risk management; ratify valuations by an independent valuer; review inter-corporate loans and investments; approve related party contracts; review the financial statements; review the auditors' independence and performance.", "f": [], "s": ["177", "139", "144", "186", "188"], "r": ["B-1.1.*"]},
         {"t": "Nomination and Remuneration Committee", "d": "Propose directors and key managerial personnel for appointment; formulate the remuneration policy; formulate and recommend employee stock option and sweat equity plans.", "f": [], "s": ["178", "54", "62"], "r": ["B-2.1.*"]},
         {"t": "Corporate Social Responsibility Committee", "d": "The Board constitutes the committee and delegates powers to it.", "f": [], "s": ["135"], "r": ["A-3.4.6"]},
         {"t": "Minute the decisions", "d": "Record the committee's recommendations in its minutes and place them before the Board.", "f": [], "s": ["118"]}],
     "check": ["Committee constituted with the required members", "Notice and agenda sent", "Quorum met", "Recommendations minuted", "Placed before the Board"],
     "rs": [{"t": "Audit Committee", "ids": ["B-1.1.*"]}, {"t": "Nomination and Remuneration Committee", "ids": ["B-2.1.*"]}],
     "drafts": ["audit committee", "remuneration committee", "nomination"],
     "lodr": [{"id": "18", "why": "Audit committee of listed companies."}, {"id": "19", "why": "Nomination and remuneration committee of listed companies."}]},
    {"id": "egm", "title": "Extraordinary General Meeting",
     "summary": "A general meeting held outside the yearly AGM, called by the board or on a requisition by members.",
     "when": "A matter needs members' approval before the next AGM.",
     "steps": [
         {"t": "Board meeting to call the EGM", "d": "Approve the resolutions, the explanatory statement and the notice, fix the date, time and place, and authorise the Company Secretary to issue the notice.", "f": [], "s": ["100", "102", "173"]},
         {"t": "Requisition by members", "d": "Members who hold the required part of the voting power may require the board to call a meeting. The board must act within the period the Act gives.", "f": [], "s": ["100"]},
         {"t": "Send the notice and explanatory statement", "d": "Send to every member, director and the auditor. The statement explains each item of special business and the interest of directors and key managerial personnel.", "f": [], "s": ["101", "102"]},
         {"t": "Hold the meeting and vote", "d": "Check quorum, proxies, and e-voting or poll as the Act requires.", "f": [], "s": ["103", "107", "108"]},
         {"t": "Minutes and filing", "d": "File the resolutions that must be filed with the Registrar.", "f": ["MGT-14"], "s": ["117", "118"]}],
     "check": ["Board approval of the notice", "Explanatory statement attached", "Notice sent in time", "Quorum and proxies checked", "Voting recorded", "Minutes signed", "MGT-14 filed"],
     "rs": [{"t": "Resolutions often put to an EGM: Memorandum and Articles", "ids": ["C-2.2.*"]}, {"t": "Share capital", "ids": ["C-2.1.*"]}, {"t": "Borrowings and charges", "ids": ["C-2.4.*"]}],
     "more": [["All general meeting resolutions", "C"]],
     "drafts": ["egm", "extraordinary", "requisition", "explanatory statement"],
     "lodr": [{"id": "44", "why": "Voting results of the meeting."}]},
    {"id": "appoint-director", "title": "Appointing a director",
     "summary": "Consent, disclosures, approval by the board and members, and the filing with the Registrar.",
     "when": "A new director is appointed, or an additional, alternate, nominee or independent director is brought in.",
     "steps": [
         {"t": "Check eligibility", "d": "Check the person's DIN, that they are not disqualified, and the limit on the number of directorships.", "f": ["DIR-8"], "s": ["164", "165"]},
         {"t": "Collect consent and disclosures", "d": "Take the director's consent to act (DIR-2), the declaration of non-disqualification and the disclosure of interest (MBP-1). The Board records them.", "f": ["DIR-2", "DIR-8", "MBP-1"], "s": ["152", "184"], "r": ["A-3.4.7", "A-1.2.5"]},
         {"t": "Board resolution", "d": "The board appoints as additional director if the Articles allow, or fills a casual vacancy, or appoints an alternate or nominee director. Where the members must appoint, the board recommends the name.", "f": [], "s": ["161", "152"], "r": ["A-3.4.7", "A-3.4.8", "A-3.4.9", "A-3.4.10"]},
         {"t": "Members' approval", "d": "Pass the resolution at a general meeting where the appointment needs it. A director who retires by rotation has a resolution of his own. Independent directors have extra conditions.", "f": ["MGT-14"], "s": ["149", "150", "152", "160"], "r": ["C-2.6.5", "C-2.6.6", "C-2.6.8"]},
         {"t": "Issue the appointment letter", "d": "Issue the letter with the terms of appointment.", "f": [], "s": ["149"]},
         {"t": "File with the Registrar", "d": "File the return of appointment within the time the Act gives.", "f": ["DIR-12"], "s": ["170"]},
         {"t": "Update the registers", "d": "Update the register of directors and the register of directors' shareholding.", "f": [], "s": ["170"]}],
     "check": ["DIN available", "Consent in the form required", "Non-disqualification declaration received", "Disclosure of interest taken", "Board resolution passed", "Members' resolution passed where needed", "DIR-12 filed in time", "Registers updated"],
     "rs": [{"t": "Board meeting", "ids": ["A-3.4.7", "A-3.4.8", "A-3.4.9", "A-3.4.10", "A-1.2.5"]}, {"t": "General meeting", "ids": ["C-2.6.5", "C-2.6.6", "C-2.6.8", "C-2.6.9"]}],
     "drafts": ["appointment of director", "appointment letter", "director consent", "consent to act", "dir-2", "dir-8"],
     "lodr": [{"id": "17", "why": "Board composition for listed companies."}, {"id": "30", "why": "Disclosure of the appointment to the exchanges."}]},
    {"id": "remove-director", "title": "Removal or resignation of a director",
     "summary": "Removal by the members on special notice, or acceptance of a director's resignation by the Board.",
     "when": "A director is to be removed, or has sent a resignation.",
     "steps": [
         {"t": "Removal: special notice by a member", "d": "A member gives special notice of a resolution to remove a director. The Board notes it and includes it in the agenda of the general meeting.", "f": [], "s": ["169", "115"], "r": ["A-3.4.13"]},
         {"t": "Send the notice and give the director a hearing", "d": "Send the notice of the general meeting with the explanatory statement and give the director concerned a chance to be heard.", "f": [], "s": ["169", "101", "102"]},
         {"t": "Pass the resolution", "d": "Pass the resolution at the general meeting.", "f": [], "s": ["169"], "r": ["C-2.6.4"]},
         {"t": "Resignation: take note", "d": "Place the resignation notice before the Board, accept it, record the director's contribution, and authorise the filing.", "f": ["DIR-12"], "s": ["168"], "r": ["A-3.4.14"]},
         {"t": "File and update records", "d": "File the return with the Registrar and update the register of directors.", "f": ["DIR-12"], "s": ["168", "170"]}],
     "check": ["Notice or resignation letter received", "Board resolution passed", "General meeting resolution passed (removal)", "DIR-12 filed in time", "Registers updated", "Exchange informed if listed"],
     "rs": [{"t": "Model resolutions", "ids": ["A-3.4.13", "A-3.4.14", "C-2.6.4"]}],
     "drafts": ["resignation", "removal of director"],
     "lodr": [{"id": "30", "why": "Resignation and removal of directors are disclosed by listed companies."}]},
    {"id": "remuneration", "title": "Director remuneration and Managing Director appointment",
     "summary": "Appointment of a managing or whole-time director and approval of remuneration within or above the limits.",
     "when": "A managing or whole-time director is appointed or paid, or remuneration crosses the limits.",
     "steps": [
         {"t": "Board resolution", "d": "The Board, on the recommendation of the Nomination and Remuneration Committee where one exists, approves the appointment and terms and the draft explanatory statement.", "f": [], "s": ["196", "197", "203", "178"], "r": ["A-3.4.5", "A-3.4.2"]},
         {"t": "Members' approval", "d": "Obtain the members' approval for the appointment and remuneration.", "f": [], "s": ["196", "197", "198"], "r": ["C-2.6.11", "C-2.6.12"]},
         {"t": "Special cases", "d": "Remuneration or commission to non-executive directors; waiver of recovery of excess remuneration; remuneration above limits in a loss-making company; appointment of a director's relative to an office of profit.", "f": [], "s": ["197", "198", "188"], "r": ["C-2.6.1", "C-2.6.2", "C-2.6.3", "C-2.6.10", "C-2.6.9", "A-3.4.11"]},
         {"t": "File the appointment", "d": "File the return of appointment of the managing director or whole-time director.", "f": ["MR-1"], "s": ["196", "203"]},
         {"t": "Key Managerial Personnel", "d": "The Board appoints key managerial personnel: MD, CEO, CFO and Company Secretary, and may delegate powers to them.", "f": [], "s": ["203"], "r": ["A-3.4.4", "A-3.4.3"]}],
     "check": ["Nomination and Remuneration Committee recommendation", "Board resolution", "Explanatory statement prepared", "Members' resolution passed", "Central Government approval if the limits need it", "MR-1 filed", "Agreement or appointment letter signed"],
     "rs": [{"t": "Board meeting", "ids": ["A-3.4.2", "A-3.4.3", "A-3.4.4", "A-3.4.5", "A-3.4.11"]}, {"t": "General meeting", "ids": ["C-2.6.1", "C-2.6.2", "C-2.6.3", "C-2.6.9", "C-2.6.10", "C-2.6.11", "C-2.6.12"]}],
     "drafts": ["managing director", "remuneration", "whole-time"],
     "lodr": [{"id": "17", "why": "Board matters for listed companies."}, {"id": "19", "why": "Nomination and remuneration committee."}]},
    {"id": "appoint-auditor", "title": "Appointing a statutory auditor",
     "summary": "Eligibility, consent, resolution and the filing with the Registrar.",
     "when": "First auditor after incorporation, at each term, or to fill a casual vacancy.",
     "steps": [
         {"t": "Check eligibility", "d": "Obtain the auditor's written consent and a certificate that the appointment meets the conditions of the Act.", "f": [], "s": ["139", "141"]},
         {"t": "First auditor", "d": "The Board appoints the first auditor, who holds office until the first AGM.", "f": [], "s": ["139"], "r": ["A-1.2.13"]},
         {"t": "Appointment or reappointment at the AGM", "d": "The Board recommends the auditor at the meeting that approves the accounts, the Audit Committee recommends where there is one, and the members appoint at the AGM for the term the Act gives. The remuneration is fixed by the members or as they authorise.", "f": [], "s": ["139", "141", "177"], "r": ["A-2.1.8", "A-3.1.2", "B-1.1.2", "C-1.1.1", "C-1.1.5"]},
         {"t": "Casual vacancy", "d": "Fill a casual vacancy as the Act provides.", "f": [], "s": ["139"], "r": ["A-3.1.3"]},
         {"t": "Intimate the auditor", "d": "Send the appointment letter to the auditor.", "f": [], "s": ["139"]},
         {"t": "File with the Registrar", "d": "File the notice of appointment within the time the Act gives.", "f": ["ADT-1"], "s": ["139"]}],
     "check": ["Auditor's consent and eligibility certificate", "Board or members' resolution", "Appointment letter sent", "ADT-1 filed in time", "Remuneration fixed by the right body"],
     "rs": [{"t": "Model resolutions", "ids": ["A-1.2.13", "A-2.1.8", "A-3.1.1", "A-3.1.2", "A-3.1.3", "A-2.1.5", "A-2.1.10", "B-1.1.2", "C-1.1.1", "C-1.1.5", "C-2.3.1"]}],
     "drafts": ["auditor", "adt"],
     "lodr": []},
    {"id": "allotment", "title": "Allotting shares",
     "summary": "Increase of capital if needed, the offer, the resolutions, the allotment and the returns.",
     "when": "The company issues new shares by rights issue, private placement or otherwise.",
     "steps": [
         {"t": "Check authorised capital and the Articles", "d": "If the authorised capital is not enough, increase it first.", "f": ["SH-7"], "s": ["61", "64"], "r": ["A-3.6.10"]},
         {"t": "Board and members' approvals", "d": "Pass the Board resolution, call a general meeting for the special resolution where the type of issue requires it, and record the valuation report where one is needed.", "f": ["MGT-14"], "s": ["62", "42", "179"], "r": ["A-3.6.9", "A-3.6.4", "A-3.6.5", "C-2.1.13", "C-2.1.5", "C-2.1.6"]},
         {"t": "Make the offer", "d": "Send the offer letter (PAS-4 for private placement) and take the application money in the way the Act and Rules require.", "f": ["PAS-4", "PAS-5"], "s": ["42", "62"]},
         {"t": "Allot the shares", "d": "Pass the allotment resolution after the money is received, and issue the share certificates in time.", "f": [], "s": ["39", "56"], "r": ["A-3.6.8", "A-3.6.3"]},
         {"t": "File the return of allotment", "d": "File the return with the Registrar. The time limit differs between types of issue, so check the section and the Rule.", "f": ["PAS-3"], "s": ["39", "42"]},
         {"t": "Update the registers", "d": "Enter the holders in the register of members and the record of private placement offers.", "f": ["PAS-5"], "s": ["88"]},
         {"t": "Later steps", "d": "Calls on partly paid shares, duplicate share certificates, forfeiture and conversion of preference shares are covered by their own resolutions.", "f": [], "s": ["49", "46", "55"], "r": ["A-3.6.11", "A-3.2.4", "A-3.2.6", "A-3.2.2"]}],
     "check": ["Authorised capital sufficient", "Board and special resolutions passed", "Offer letter sent", "Application money received in the bank account", "Allotment resolution", "Share certificates issued", "PAS-3 filed", "Registers updated", "Valuation report where required"],
     "rs": [{"t": "Fresh issues: Board meeting", "ids": ["A-3.6.*"]}, {"t": "Fresh issues: general meeting", "ids": ["C-2.1.4", "C-2.1.5", "C-2.1.6", "C-2.1.13"]}],
     "drafts": ["allotment", "share certificate", "rights issue", "private placement", "offer letter"],
     "lodr": [{"id": "30", "why": "Listed companies disclose the issue and allotment."}]},
    {"id": "bonus-split", "title": "Bonus shares, stock split and buy-back",
     "summary": "Changes to share capital that do not raise fresh money: bonus issue, split of shares and buy-back.",
     "when": "The company issues bonus shares, splits its shares or buys back its own shares.",
     "steps": [
         {"t": "Check the Articles", "d": "A bonus issue needs authority in the Articles, so the Articles may have to be amended first.", "f": [], "s": ["63", "14"], "r": ["C-2.1.3"]},
         {"t": "Bonus issue: Board and members", "d": "The Board recommends capitalising free reserves. The members approve at a general meeting. Then shares are allotted and the return of allotment is filed.", "f": ["MGT-14", "PAS-3"], "s": ["63"], "r": ["A-3.2.7", "A-3.2.8", "C-2.1.8"]},
         {"t": "Stock split or consolidation", "d": "Alter the capital clause and the Articles to change the face value, as the Board and the members approve.", "f": ["SH-7"], "s": ["61"], "r": ["A-3.2.9", "A-3.2.10", "C-2.1.10"]},
         {"t": "Buy-back", "d": "The Board approves the buy-back, within the percentage of paid-up capital and free reserves the Act allows, and the members approve by special resolution where needed. The explanatory statement carries the facts about promoters' holdings. Make the declaration of solvency and the letter of offer, buy back within the period the Rules give, and extinguish the shares.", "f": ["SH-8", "SH-9", "SH-11", "SH-15"], "s": ["68", "69", "70"], "r": ["A-3.10.3", "C-2.1.7"]},
         {"t": "Update the records", "d": "Update the register of members and the capital in the records, and file the returns.", "f": [], "s": ["88"]}],
     "check": ["Articles allow the issue", "Board resolution passed", "General meeting resolution passed", "Record date fixed", "Allotment made and PAS-3 filed", "Buy-back limits worked out", "Declaration of solvency filed (buy-back)", "Shares extinguished (buy-back)", "Registers updated"],
     "rs": [{"t": "Bonus, split and buy-back", "ids": ["A-3.2.7", "A-3.2.8", "A-3.2.9", "A-3.2.10", "A-3.10.3", "C-2.1.3", "C-2.1.7", "C-2.1.8", "C-2.1.10"]}],
     "drafts": ["bonus", "buy-back", "buyback", "split"],
     "lodr": [{"id": "30", "why": "Listed companies disclose these corporate actions."}]},
    {"id": "esop", "title": "ESOP and sweat equity",
     "summary": "Employee stock option plans and sweat equity: committee recommendation, Board, special resolution.",
     "when": "The company wants to grant stock options or issue sweat equity shares.",
     "steps": [
         {"t": "Committee recommends", "d": "The Nomination and Remuneration Committee formulates the scheme and recommends the employees.", "f": [], "s": ["178"], "r": ["B-2.1.1", "B-2.1.2", "B-2.1.3", "B-2.1.7"]},
         {"t": "Board approves", "d": "The Board approves the scheme, or the issue of sweat equity based on a registered valuer's valuation.", "f": [], "s": ["62", "54"], "r": ["A-3.6.1", "A-3.6.6", "A-3.6.7"]},
         {"t": "Members approve", "d": "Pass a special resolution with an explanatory statement.", "f": ["MGT-14"], "s": ["62", "54"], "r": ["C-2.3.2", "C-2.1.9"]},
         {"t": "Issue and allot", "d": "Grant and allot as the scheme and the Rules provide, and file the return of allotment.", "f": ["PAS-3"], "s": ["39", "54", "62"]}],
     "check": ["Committee recommendation", "Valuation report (sweat equity)", "Board resolution", "Special resolution passed", "MGT-14 filed", "Allotment made, PAS-3 filed"],
     "rs": [{"t": "Model resolutions", "ids": ["B-2.1.*", "A-3.6.1", "A-3.6.6", "A-3.6.7", "C-2.1.9", "C-2.3.2"]}],
     "drafts": ["esop", "stock option", "sweat equity"],
     "lodr": []},
    {"id": "dividend", "title": "Declaring and paying a dividend",
     "summary": "Interim and final dividend, the bank account for it, and transfer of unpaid dividend.",
     "when": "The company recommends or declares a dividend.",
     "steps": [
         {"t": "Board recommends or declares", "d": "On the audited accounts, the Board recommends a final dividend for approval at the AGM, or declares an interim dividend. The Board also approves closure of the register of members for it.", "f": [], "s": ["123", "91"], "r": ["A-3.10.4", "A-3.10.5", "A-3.10.1", "A-2.1.3", "A-2.1.6"]},
         {"t": "Members declare the final dividend", "d": "The AGM declares the dividend, which cannot exceed what the Board recommended.", "f": [], "s": ["123"], "r": ["C-1.1.2", "C-2.1.1"]},
         {"t": "Open a separate bank account", "d": "Open a bank account for the dividend and deposit the amount within the days the Act gives after the declaration, as the model resolution states.", "f": [], "s": ["123"], "r": ["A-3.10.2"]},
         {"t": "Pay the members", "d": "Pay or post the dividend within the period the Act gives.", "f": [], "s": ["127"]},
         {"t": "Unpaid dividend", "d": "Dividend that stays unpaid goes to the Unpaid Dividend Account. Amounts unpaid for seven years go to the Investor Education and Protection Fund.", "f": [], "s": ["124", "125"], "r": ["A-3.10.6"]}],
     "check": ["Audited accounts approved", "Board resolution on dividend", "Register of members closed or record date fixed", "AGM resolution (final dividend)", "Dividend bank account opened", "Dividend paid within the time allowed", "Unpaid dividend account maintained", "Transfer to IEPF on the due date"],
     "rs": [{"t": "Model resolutions", "ids": ["A-3.10.1", "A-3.10.2", "A-3.10.4", "A-3.10.5", "A-3.10.6", "A-2.1.6", "C-1.1.2", "C-2.1.1"]}],
     "drafts": ["dividend"],
     "lodr": [{"id": "42", "why": "Record date and notice to the exchange for listed companies."}]},
    {"id": "borrowing", "title": "Borrowing, loans, investments and deposits",
     "summary": "Board and members' approval of borrowings, security, loans, investments and guarantees, and of deposits.",
     "when": "The company borrows, creates security, lends, invests, gives a guarantee or accepts deposits.",
     "steps": [
         {"t": "Board approval", "d": "Board resolution for the borrowing, from a bank or institution, from directors or key managerial personnel, as inter-corporate deposits, or by issuing debentures.", "f": [], "s": ["179", "71"], "r": ["A-3.11.3", "A-3.11.4", "A-3.11.5", "A-3.11.6", "A-3.11.10", "A-3.11.11"]},
         {"t": "Members' approval where limits are crossed", "d": "Borrowing above paid-up capital and free reserves, and creation of a charge over the undertaking, need a special resolution.", "f": ["MGT-14"], "s": ["180"], "r": ["A-3.11.9", "C-2.4.1", "C-2.4.4", "C-2.4.5"]},
         {"t": "Loans, investments and guarantees", "d": "Approval of loans, investments, securities and guarantees against the limits of the Act. Loans to directors have their own rule.", "f": [], "s": ["185", "186", "187"], "r": ["A-3.11.1", "A-3.4.1", "B-1.1.6", "C-2.7.2", "C-2.7.3", "C-2.7.4"]},
         {"t": "Deposits", "d": "Accepting deposits from members or the public follows the Deposits Rules: a circular or advertisement, and returns.", "f": ["DPT-1", "DPT-3"], "s": ["73", "76"], "r": ["A-3.11.7", "A-3.11.8", "C-2.4.2"]},
         {"t": "Register the charge", "d": "File the particulars of any charge created.", "f": ["CHG-1"], "s": ["77", "78"]}],
     "check": ["Board resolution", "Members' resolution where limits are crossed", "Limits under sections 180 and 186 checked", "Charge documents signed", "CHG-1 filed", "Deposit circular and returns filed", "Registers updated"],
     "rs": [{"t": "Board meeting", "ids": ["A-3.11.*"]}, {"t": "General meeting", "ids": ["C-2.4.*", "C-2.7.*"]}],
     "drafts": ["borrow", "loan", "guarantee", "investment", "debenture", "deposit"],
     "lodr": [{"id": "23", "why": "Related party transactions of listed companies."}]},
    {"id": "charge", "title": "Creating and satisfying a charge",
     "summary": "Registering a charge on the company's assets with the Registrar, and filing its satisfaction.",
     "when": "The company takes a loan secured on its assets, or repays one.",
     "steps": [
         {"t": "Board approval", "d": "Approve the borrowing and the security. Check whether members' consent is needed for the limit.", "f": [], "s": ["179", "180"], "r": ["A-3.11.6"]},
         {"t": "Members' special resolution", "d": "A mortgage or charge over the whole or substantially the whole undertaking needs the consent of the members.", "f": ["MGT-14"], "s": ["180"], "r": ["C-2.4.1"]},
         {"t": "Execute the documents", "d": "Sign the loan and security documents with the lender.", "f": [], "s": ["77"]},
         {"t": "Register the charge", "d": "File the particulars of the charge within the time the Act gives. Late filing is possible on extra fees within a longer period.", "f": ["CHG-1"], "s": ["77", "78"]},
         {"t": "Keep the register of charges", "d": "Enter the charge in the company's register.", "f": [], "s": ["85"]},
         {"t": "On repayment, file satisfaction", "d": "File the satisfaction of the charge once the loan is repaid.", "f": ["CHG-4"], "s": ["82"]}],
     "check": ["Board resolution", "Members' resolution where required", "Charge documents signed", "CHG-1 filed in time", "Register of charges updated", "CHG-4 filed on repayment"],
     "rs": [{"t": "Model resolutions", "ids": ["A-3.11.6", "C-2.4.1"]}],
     "drafts": ["creation of charge", "mortgage", "hypothecation"],
     "lodr": []},
    {"id": "annual-filing", "title": "Yearly filings after the AGM",
     "summary": "The returns a company files every year with the Registrar.",
     "when": "Each financial year, after the AGM.",
     "steps": [
         {"t": "Authorise the filings", "d": "The Board authorises directors or the Company Secretary to file e-forms and records with the Registrar.", "f": [], "s": ["92", "137"], "r": ["A-2.1.2"]},
         {"t": "File the financial statements", "d": "File with the Registrar after the AGM, within the time the Act gives.", "f": ["AOC-4"], "s": ["137"]},
         {"t": "File the annual return", "d": "File the annual return of members, directors and meetings.", "f": ["MGT-7"], "s": ["92"]},
         {"t": "File the auditor's appointment if due", "d": "If a new term begins, file the notice.", "f": ["ADT-1"], "s": ["139"]},
         {"t": "Return of deposits and exempt receipts", "d": "File the return showing money received that is not a deposit, by its due date.", "f": ["DPT-3"], "s": ["73"]},
         {"t": "Director KYC", "d": "Every director with a DIN files the yearly KYC by the due date.", "f": ["DIR-3"], "s": ["153"]},
         {"t": "Half-yearly return of dues to small enterprises", "d": "Companies that owe such dues beyond the stated period file the half-yearly return.", "f": ["MSME-1"], "s": ["405"]},
         {"t": "Compliance status", "d": "Place a periodic compliance status report before the Board.", "f": [], "s": ["134"], "r": ["A-2.1.1"]}],
     "check": ["AOC-4 filed", "MGT-7 filed", "DPT-3 filed", "Director KYC done", "MSME-1 half-yearly filed if applicable", "Register of members, directors and charges updated", "Compliance status report taken on record"],
     "rs": [{"t": "Model resolutions", "ids": ["A-2.1.1", "A-2.1.2", "C-1.1.3"]}],
     "drafts": ["annual return", "financial statements", "aoc", "mgt"],
     "lodr": []},
    {"id": "change-name", "title": "Changing the company's name",
     "summary": "Name approval, the resolution and the Registrar's fresh certificate.",
     "when": "The company wants a new name.",
     "steps": [
         {"t": "Reserve the new name", "d": "Apply for the new name with the Registrar. The Board resolution lists names in order of preference.", "f": [], "s": ["4", "13"], "r": ["A-3.9.15"]},
         {"t": "Board and members' approval", "d": "Pass a Board resolution and a special resolution, and alter the Memorandum and Articles.", "f": ["MGT-14"], "s": ["13", "14"], "r": ["C-2.2.3"]},
         {"t": "File the application", "d": "File the application for change of name with the Registrar.", "f": ["INC-24"], "s": ["13"]},
         {"t": "Receive the fresh certificate", "d": "The Registrar issues a fresh certificate. Update PAN, GST, bank, licences and letterheads.", "f": [], "s": ["13"]}],
     "check": ["Name reserved", "Special resolution passed", "MGT-14 filed", "INC-24 filed", "Fresh certificate received", "Authorities informed"],
     "rs": [{"t": "Model resolutions", "ids": ["A-3.9.15", "C-2.2.3"]}],
     "drafts": ["change of name", "name change"],
     "lodr": [{"id": "30", "why": "Listed companies inform the exchanges."}]},
    {"id": "alter-moa-aoa", "title": "Altering the Memorandum or Articles; changing the type of company",
     "summary": "A special resolution, a filing, and some changes that need extra approval, including conversion between private and public.",
     "when": "The company changes its objects, capital clause or its Articles, or converts from private to public or the reverse.",
     "steps": [
         {"t": "Board resolution", "d": "Approve the alteration and the notice for the members' meeting.", "f": [], "s": ["13", "14", "173"], "r": ["A-3.9.5"]},
         {"t": "Special resolution", "d": "Pass it at a general meeting. A change of objects may need a postal ballot.", "f": ["MGT-14"], "s": ["13", "14", "110"], "r": ["C-2.2.1", "C-2.2.2"]},
         {"t": "Private to public, or public to private", "d": "The Articles and the name change, and the conversion is approved as the Act provides.", "f": ["INC-27", "MGT-14"], "s": ["14", "2"], "r": ["A-3.9.12", "A-3.9.2"]},
         {"t": "File the altered documents", "d": "File the resolution and the altered documents with the Registrar within the time the Act gives.", "f": ["MGT-14"], "s": ["14", "117"]},
         {"t": "Registrar's confirmation where needed", "d": "Where the Act needs the Registrar or the Regional Director to confirm, wait for the order before acting on the change.", "f": ["INC-28"], "s": ["13"]}],
     "check": ["Board approval", "Special resolution passed", "Altered documents prepared", "MGT-14 filed", "Order received where required"],
     "rs": [{"t": "Model resolutions", "ids": ["A-3.9.2", "A-3.9.5", "A-3.9.12", "C-2.2.1", "C-2.2.2"]}],
     "drafts": ["alteration of memorandum", "alteration of articles", "amendment of articles"],
     "lodr": []},
    {"id": "authorised-capital", "title": "Increasing authorised capital",
     "summary": "Altering the capital clause and filing the notice.",
     "when": "The company wants authorised capital above its present limit.",
     "steps": [
         {"t": "Check the Articles", "d": "Confirm the Articles authorise an increase. If they do not, alter them first.", "f": [], "s": ["61", "14"], "r": ["C-2.1.11"]},
         {"t": "Board resolution", "d": "Approve the increase and call the general meeting.", "f": [], "s": ["61", "173"], "r": ["A-3.6.10"]},
         {"t": "Members' resolution", "d": "Pass the resolution at the meeting and alter the Memorandum. Reclassification of capital has its own resolution.", "f": ["MGT-14"], "s": ["61", "117"], "r": ["C-2.1.12", "C-2.1.4"]},
         {"t": "File the notice and pay stamp duty", "d": "File the notice of the increase and the altered Memorandum.", "f": ["SH-7"], "s": ["64"]}],
     "check": ["Articles checked", "Board and members' resolutions", "SH-7 filed in time", "Stamp duty and fees paid", "Altered Memorandum filed"],
     "rs": [{"t": "Model resolutions", "ids": ["A-3.6.10", "C-2.1.4", "C-2.1.11", "C-2.1.12"]}],
     "drafts": ["authorised capital", "authorized capital", "increase of capital"],
     "lodr": []},
    {"id": "related-party", "title": "Related party transactions",
     "summary": "Approval of the audit committee, the board and, where needed, the members.",
     "when": "The company proposes a transaction with a related party.",
     "steps": [
         {"t": "Identify the related party and the transaction", "d": "Check the list of related parties against the transaction.", "f": [], "s": ["2", "188"]},
         {"t": "Audit committee and board approval", "d": "Obtain approval, with interested directors keeping out of the vote.", "f": ["MBP-1"], "s": ["177", "184", "188"], "r": ["B-1.1.7", "A-3.4.12"]},
         {"t": "Members' approval if the limits are crossed", "d": "Pass the resolution by the members who are not related. The explanatory statement gives the facts of the transaction.", "f": ["MGT-14"], "s": ["188"], "r": ["C-2.6.7", "C-2.6.9", "C-2.5.4"]},
         {"t": "Disclose", "d": "Record the transaction in the board's report and the register of contracts.", "f": ["MBP-4"], "s": ["189", "134"]}],
     "check": ["Related party list updated", "Arm's length and ordinary course tested", "Audit committee approval", "Board approval", "Members' approval if needed", "Register of contracts updated", "Disclosure in the board's report"],
     "rs": [{"t": "Model resolutions", "ids": ["B-1.1.7", "A-3.4.12", "A-3.4.11", "C-2.6.7", "C-2.6.9"]}],
     "drafts": ["related party"],
     "lodr": [{"id": "23", "why": "Related party transactions of listed companies."}]},
    {"id": "restructuring", "title": "Merger, sale and lease of an undertaking, spin-off",
     "summary": "Business restructuring: schemes of arrangement, sale or lease of an undertaking, and spinning off a business.",
     "when": "The company merges, sells or leases an undertaking, or demerges a business.",
     "steps": [
         {"t": "Board approves the proposal", "d": "A detailed business plan and the draft scheme go to the Board. For a merger the Board also approves the application under section 230. For a sale or lease, the Board recommends it to the members.", "f": [], "s": ["230", "232", "180"], "r": ["A-3.3.1", "A-3.3.4", "A-3.3.8", "A-3.3.9", "A-3.3.2", "A-3.3.3", "A-3.3.10"]},
         {"t": "Members' approval", "d": "A special resolution with a statement of material facts: profile of the companies, salient features of the scheme, share swap ratio and the shareholding before and after.", "f": ["MGT-14"], "s": ["180", "230", "102"], "r": ["C-2.5.1", "C-2.5.2", "C-2.5.3", "C-2.5.4"]},
         {"t": "Approval of the Tribunal (schemes)", "d": "A scheme of arrangement or merger is sanctioned by the Tribunal. File its order with the Registrar.", "f": ["INC-28"], "s": ["230", "231", "232"]},
         {"t": "Give effect", "d": "Pass the Board resolutions that implement the scheme and update the records.", "f": [], "s": ["232"]}],
     "check": ["Valuation and fairness opinion", "Board resolution", "Special resolution passed", "Tribunal application filed and order received", "Order filed with the Registrar", "Records updated"],
     "rs": [{"t": "Board meeting", "ids": ["A-3.3.*"]}, {"t": "General meeting", "ids": ["C-2.5.*"]}],
     "drafts": ["merger", "amalgamation", "demerger", "sale of undertaking", "lease"],
     "lodr": [{"id": "37", "why": "Schemes of arrangement of listed entities."}, {"id": "30", "why": "Disclosure to the exchanges."}]},
    {"id": "strike-off", "title": "Striking off a company's name",
     "summary": "Applying to the Registrar to remove the name of a company that has no business.",
     "when": "The company has not carried on business for the period the Act gives and has no liabilities.",
     "steps": [
         {"t": "Check that the company can apply", "d": "Confirm the conditions and prepare the statement of accounts.", "f": [], "s": ["248"]},
         {"t": "Pass resolutions", "d": "Pass a special resolution or obtain the consent of the required part of members by shares.", "f": ["MGT-14"], "s": ["248"]},
         {"t": "Prepare the documents", "d": "Affidavits by the directors, the indemnity bond, and the statement of accounts.", "f": [], "s": ["248"]},
         {"t": "File the application", "d": "File the application with the Registrar.", "f": ["STK-2"], "s": ["248"]},
         {"t": "Voluntary winding up", "d": "A solvent company that wants to wind up instead follows a separate route with a special resolution.", "f": [], "s": ["304", "305"], "r": ["A-3.9.4", "C-2.8.1"]}],
     "check": ["No business for the required period", "No assets or liabilities", "Special resolution passed", "Affidavits and indemnity bond signed", "STK-2 filed"],
     "rs": [{"t": "Voluntary winding up", "ids": ["A-3.9.4", "C-2.8.1"]}],
     "drafts": ["strike off", "dissolution", "indemnity"],
     "lodr": []},
]


def words(s):
    return re.findall(r"[a-z0-9]+", s.lower())


def drafts_list():
    p = SITE / "drafts" / "data" / "index.json"
    try:
        return json.loads(p.read_text(encoding="utf-8")).get("drafts", [])
    except (OSError, ValueError):
        return []


def load_res():
    p = SITE / "data" / "res" / "index.json"
    try:
        return json.loads(p.read_text(encoding="utf-8")).get("items", [])
    except (OSError, ValueError):
        return []


def expand(pats, items):
    import fnmatch
    out = []
    for pat in pats:
        for it in items:
            if fnmatch.fnmatchcase(it["id"], pat) and it["id"] not in out:
                out.append(it["id"])
    return out


def main():
    procs = list(PROCESSES)
    extra = Path("guide") / "processes.json"
    if extra.exists():
        try:
            procs += json.loads(extra.read_text(encoding="utf-8"))
        except ValueError as e:
            print(f"guide/processes.json could not be read: {e}")
    drafts = drafts_list()
    items = load_res()
    title = {it["id"]: it["t"] for it in items}
    out = []
    for p in procs:
        steps = []
        for st in p["steps"]:
            st = dict(st)
            st["r"] = [[i, title[i]] for i in expand(st.get("r", []), items) if i in title][:16]
            steps.append(st)
        p = dict(p)
        p["steps"] = steps
        p["rs"] = [{"t": g["t"], "r": [[i, title[i]] for i in expand(g["ids"], items) if i in title]} for g in p.get("rs", [])]
        p["rs"] = [g for g in p["rs"] if g["r"]]
        hit = []
        for d in drafts:
            hay = " ".join(words(d["title"] + " " + d.get("category", "")))
            if any(k.lower() in hay for k in p.get("drafts", [])):
                hit.append({"id": d["id"], "title": d["title"], "category": d.get("category", "")})
        p["drafts"] = hit[:12]
        out.append(p)
    SITE.mkdir(exist_ok=True)
    (SITE / "data").mkdir(parents=True, exist_ok=True)
    (SITE / "data" / "processes.json").write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    idx = SITE / "data" / "index.json"
    try:
        d = json.loads(idx.read_text(encoding="utf-8"))
        d["proc"] = len(out)
        idx.write_text(json.dumps(d, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    except (OSError, ValueError):
        pass
    print(f"{len(out)} processes written; {sum(len(p['drafts']) for p in out)} draft matches")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Processes step ended with an error: {e}")
    sys.exit(0)
