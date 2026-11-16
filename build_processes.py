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
     "summary": "From name approval to the certificate of incorporation, and the filings that follow in the first months.",
     "when": "Starting a new private company through the SPICe+ route.",
     "steps": [
         {"t": "Get digital signatures and DINs", "d": "Every proposed director and subscriber needs a Digital Signature Certificate. DIN is applied for inside the incorporation form for new directors.", "f": ["DIR-3"], "s": ["153", "154"]},
         {"t": "Choose and reserve the name", "d": "Check the name against the Registrar's name rules and trade marks. The name is applied for in Part A of the incorporation form.", "f": ["INC-32"], "s": ["4"]},
         {"t": "Settle the memorandum and articles", "d": "Prepare the Memorandum (objects, liability, capital) and Articles. Use the ready-made forms for the standard version or file your own.", "f": ["INC-33", "INC-34"], "s": ["3", "4", "5"]},
         {"t": "File the incorporation form", "d": "File SPICe+ with its attachments (proof of registered office, director and subscriber documents, declarations) and pay stamp duty and fees.", "f": ["INC-32", "INC-35"], "s": ["7"]},
         {"t": "Receive the certificate of incorporation", "d": "The Registrar issues the certificate with the CIN. PAN and TAN come with it. The date on it is the date of incorporation.", "f": [], "s": ["7", "9"]},
         {"t": "Hold the first board meeting", "d": "Appoint the first auditor, take disclosures of interest from directors, open the bank account, approve share certificates and adopt the common seal if used.", "f": ["MBP-1", "DIR-8"], "s": ["173", "184", "139"]},
         {"t": "Confirm the registered office", "d": "File the notice of the registered office address.", "f": ["INC-22"], "s": ["12"]},
         {"t": "Issue share certificates to subscribers", "d": "Deliver share certificates to the subscribers within the period the Act gives after incorporation.", "f": [], "s": ["56"]},
         {"t": "File the auditor appointment", "d": "Inform the Registrar of the auditor's appointment.", "f": ["ADT-1"], "s": ["139"]},
         {"t": "File the declaration for commencement of business", "d": "Directors file a declaration that every subscriber has paid for the shares agreed to be taken.", "f": ["INC-20A"], "s": ["10A"]}],
     "check": ["Digital signatures of all directors and subscribers", "Name available and not identical or too like another", "Draft Memorandum and Articles approved by subscribers", "Proof of registered office address and owner's no-objection", "Identity and address proofs of directors and subscribers", "Stamp duty and fees paid", "Certificate of incorporation received", "First board meeting held and minuted", "Auditor appointed and ADT-1 filed", "INC-20A filed", "Share certificates issued", "Statutory registers opened"],
     "drafts": ["incorporat", "first board", "subscriber"],
     "lodr": []},
    {"id": "registered-office", "title": "Shifting the registered office",
     "summary": "The route depends on where the office moves: within the same city, to another place in the same State, or to another State.",
     "when": "The company wants to change the address or location of its registered office.",
     "steps": [
         {"t": "Decide which case applies", "d": "Within the same city or town, within the same State but outside the city, or from one State to another. Each has different approvals.", "f": [], "s": ["12", "13"]},
         {"t": "Within the same city, town or village", "d": "A board resolution is enough. File the notice of the new address after the shift.", "f": ["INC-22"], "s": ["12"]},
         {"t": "Within the same State but under another Registrar", "d": "Pass a special resolution and obtain the Regional Director's confirmation before the shift.", "f": ["MGT-14", "INC-23"], "s": ["12", "13"]},
         {"t": "From one State to another", "d": "Amend the Memorandum by special resolution, advertise in the newspapers, notify creditors, and apply to the Regional Director for confirmation.", "f": ["MGT-14", "INC-23"], "s": ["13", "110"]},
         {"t": "Obtain the order and register it", "d": "File the Regional Director's order with both Registrars and obtain the new certificate where the State changes.", "f": ["INC-28"], "s": ["13"]},
         {"t": "Update everywhere else", "d": "Change the address with PAN, GST, bank, licences and stock exchanges, and on stationery and the website.", "f": ["INC-22"], "s": ["12"]}],
     "check": ["Board meeting called and resolution passed", "Special resolution passed if the case needs it", "Latest list of creditors and members prepared", "Newspaper advertisements published", "Application to the Regional Director filed", "Order obtained and filed", "Notice of the new address filed", "Other authorities informed"],
     "drafts": ["registered office", "shifting", "change of address"],
     "lodr": [{"id": "30", "why": "Listed companies disclose this event to the stock exchanges."}]},
    {"id": "agm", "title": "Annual General Meeting",
     "summary": "From the board meeting that fixes the date to the returns filed after the meeting.",
     "when": "Every year. The Act sets the time within which the meeting must be held.",
     "steps": [
         {"t": "Hold the board meeting", "d": "Approve the financial statements and the board's report, fix the date, time and place, and approve the notice and the explanatory statements.", "f": [], "s": ["96", "134", "173"]},
         {"t": "Appoint the scrutiniser and fix the record date", "d": "If e-voting applies, appoint a scrutiniser and fix the cut-off date for voting.", "f": [], "s": ["108"]},
         {"t": "Send the notice", "d": "Send the notice with the annual report to members, directors and the auditor, within the notice period the Act requires. Short notice needs the consent required by the Act.", "f": [], "s": ["101", "136"]},
         {"t": "Hold the meeting", "d": "Check quorum, take the chair, deal with the items, and record proxies and voting.", "f": [], "s": ["103", "105", "107"]},
         {"t": "Prepare the minutes and the scrutiniser's report", "d": "Record the proceedings and the results in the minutes book.", "f": [], "s": ["118"]},
         {"t": "File the special resolutions", "d": "File any resolution that must be filed with the Registrar.", "f": ["MGT-14"], "s": ["117"]},
         {"t": "File the financial statements and the annual return", "d": "File the financial statements and the annual return within the time the Act gives after the meeting.", "f": ["AOC-4", "MGT-7"], "s": ["137", "92"]}],
     "check": ["Financial statements signed by the board and auditor", "Board's report approved", "Notice and annual report sent in time", "Attendance register and proxies kept", "Quorum recorded", "E-voting done where it applies", "Minutes signed", "AOC-4 filed", "MGT-7 filed", "Auditor appointment or reappointment filed, if due"],
     "drafts": ["agm", "annual general", "notice of meeting", "board report", "directors report", "proxy", "attendance"],
     "lodr": [{"id": "29", "why": "Prior intimation of board meetings that consider the accounts."}, {"id": "30", "why": "Disclosure of events and the outcome of meetings."}, {"id": "44", "why": "Voting results of the meeting."}]},
    {"id": "board-meeting", "title": "Board meeting",
     "summary": "Notice, agenda, quorum, minutes and the filings that follow from the decisions.",
     "when": "Each time the board meets. The Act requires a minimum number of meetings a year with a maximum gap between two.",
     "steps": [
         {"t": "Fix the date and prepare the agenda", "d": "Prepare the agenda with the notes for each item. Check the previous minutes for matters that carry over.", "f": [], "s": ["173"]},
         {"t": "Send the notice", "d": "Send written notice, with agenda and notes, to every director within the notice period the Act requires. Hold it at shorter notice only as the Act allows.", "f": [], "s": ["173"]},
         {"t": "Check quorum and disclosures", "d": "Confirm the quorum. Take any disclosure of interest and keep interested directors out of the vote.", "f": ["MBP-1"], "s": ["174", "184"]},
         {"t": "Conduct the meeting", "d": "Take the chair, pass the resolutions item by item, and record dissent where a director asks.", "f": [], "s": ["175"]},
         {"t": "Write and sign the minutes", "d": "Circulate the draft, collect comments, and have the chairman sign within the time the Standards allow. Paste them in the minutes book.", "f": [], "s": ["118"]},
         {"t": "Act on the decisions", "d": "Make the filings that follow, for example resolutions under section 179(3), director changes or allotments.", "f": ["MGT-14", "DIR-12", "PAS-3"], "s": ["117", "179"]}],
     "check": ["Notice and agenda sent in time", "Quorum met", "Disclosures of interest taken", "Resolutions passed and recorded", "Minutes circulated and signed", "Follow-up filings made in time", "Attendance register signed"],
     "drafts": ["board meeting", "board resolution", "minutes", "agenda", "notice of board"],
     "lodr": [{"id": "29", "why": "Prior intimation to the exchange for certain agenda items."}, {"id": "30", "why": "Disclose the outcome within the time the Regulations give."}]},
    {"id": "egm", "title": "Extraordinary General Meeting",
     "summary": "A general meeting held outside the yearly AGM, called by the board or on a requisition by members.",
     "when": "A matter needs members' approval before the next AGM.",
     "steps": [
         {"t": "Board meeting to call the EGM", "d": "Approve the resolutions, the explanatory statement and the notice.", "f": [], "s": ["100", "102", "173"]},
         {"t": "Requisition by members", "d": "Members who hold the required part of the voting power may require the board to call a meeting. The board must act within the period the Act gives.", "f": [], "s": ["100"]},
         {"t": "Send the notice and explanatory statement", "d": "Send to every member, director and the auditor. The statement explains each item of special business.", "f": [], "s": ["101", "102"]},
         {"t": "Hold the meeting and vote", "d": "Check quorum, proxies, and e-voting or poll as the Act requires.", "f": [], "s": ["103", "107", "108"]},
         {"t": "Minutes and filing", "d": "File the resolutions that must be filed with the Registrar.", "f": ["MGT-14"], "s": ["117", "118"]}],
     "check": ["Board approval of the notice", "Explanatory statement attached", "Notice sent in time", "Quorum and proxies checked", "Voting recorded", "Minutes signed", "MGT-14 filed"],
     "drafts": ["egm", "extraordinary", "requisition", "explanatory statement"],
     "lodr": [{"id": "44", "why": "Voting results of the meeting."}]},
    {"id": "appoint-director", "title": "Appointing a director",
     "summary": "Consent, disclosures, approval by the board and members, and the filing with the Registrar.",
     "when": "A new director is appointed, or an additional director is appointed to fill a gap.",
     "steps": [
         {"t": "Check eligibility", "d": "Check the person's DIN, that they are not disqualified, and the limit on the number of directorships.", "f": ["DIR-8"], "s": ["164", "165"]},
         {"t": "Collect consent and disclosures", "d": "Obtain the director's consent to act, the declaration of non-disqualification and the disclosure of interest.", "f": ["DIR-2", "DIR-8", "MBP-1"], "s": ["152", "184"]},
         {"t": "Board resolution", "d": "The board appoints as additional director if the Articles allow, or recommends the name to members.", "f": [], "s": ["161", "152"]},
         {"t": "Members' approval", "d": "Pass the resolution at a general meeting where the appointment needs it. Independent directors have extra conditions.", "f": ["MGT-14"], "s": ["149", "152"]},
         {"t": "Issue the appointment letter", "d": "Issue the letter with the terms of appointment.", "f": [], "s": ["149"]},
         {"t": "File with the Registrar", "d": "File the return of appointment within the time the Act gives.", "f": ["DIR-12"], "s": ["170"]},
         {"t": "Update the registers", "d": "Update the register of directors and the register of directors' shareholding.", "f": [], "s": ["170"]}],
     "check": ["DIN available", "Consent in the form required", "Non-disqualification declaration received", "Disclosure of interest taken", "Board resolution passed", "Members' resolution passed where needed", "DIR-12 filed in time", "Registers updated"],
     "drafts": ["appointment of director", "appointment letter", "director consent", "consent to act", "dir-2", "dir-8"],
     "lodr": [{"id": "17", "why": "Board composition for listed companies."}, {"id": "30", "why": "Disclosure of the appointment to the exchanges."}]},
    {"id": "appoint-auditor", "title": "Appointing a statutory auditor",
     "summary": "Eligibility, consent, resolution and the filing with the Registrar.",
     "when": "First auditor after incorporation, at each term, or to fill a casual vacancy.",
     "steps": [
         {"t": "Check eligibility", "d": "Obtain the auditor's written consent and a certificate that the appointment meets the conditions of the Act.", "f": [], "s": ["139", "141"]},
         {"t": "Pass the resolution", "d": "The board appoints the first auditor. Later appointments are made by members at the AGM for the term the Act gives.", "f": [], "s": ["139"]},
         {"t": "Intimate the auditor", "d": "Send the appointment letter to the auditor.", "f": [], "s": ["139"]},
         {"t": "File with the Registrar", "d": "File the notice of appointment within the time the Act gives.", "f": ["ADT-1"], "s": ["139"]}],
     "check": ["Auditor's consent and eligibility certificate", "Board or members' resolution", "Appointment letter sent", "ADT-1 filed in time", "Remuneration fixed by the right body"],
     "drafts": ["auditor", "adt"],
     "lodr": []},
    {"id": "allotment", "title": "Allotting shares",
     "summary": "Increase of capital if needed, the offer, the resolutions, the allotment and the returns.",
     "when": "The company issues new shares by rights issue, private placement or otherwise.",
     "steps": [
         {"t": "Check authorised capital and the Articles", "d": "If the authorised capital is not enough, increase it first.", "f": ["SH-7"], "s": ["61", "64"]},
         {"t": "Board and members' approvals", "d": "Pass the board resolution and the special resolution where the type of issue requires it.", "f": ["MGT-14"], "s": ["62", "42", "179"]},
         {"t": "Make the offer", "d": "Send the offer letter and take the application money in the way the Act and Rules require.", "f": ["PAS-4", "PAS-5"], "s": ["42", "62"]},
         {"t": "Allot the shares", "d": "Pass the allotment resolution after the money is received, and issue the share certificates in time.", "f": [], "s": ["39", "56"]},
         {"t": "File the return of allotment", "d": "File the return with the Registrar. The time limit differs between types of issue, so check the section and the Rule.", "f": ["PAS-3"], "s": ["39", "42"]},
         {"t": "Update the registers", "d": "Enter the holders in the register of members and the record of private placement offers.", "f": ["PAS-5"], "s": ["88"]}],
     "check": ["Authorised capital sufficient", "Board and special resolutions passed", "Offer letter sent", "Application money received in the bank account", "Allotment resolution", "Share certificates issued", "PAS-3 filed", "Registers updated", "Valuation report where required"],
     "drafts": ["allotment", "share certificate", "rights issue", "private placement", "offer letter"],
     "lodr": [{"id": "30", "why": "Listed companies disclose the issue and allotment."}]},
    {"id": "charge", "title": "Creating and satisfying a charge",
     "summary": "Registering a charge on the company's assets with the Registrar, and filing its satisfaction.",
     "when": "The company takes a loan secured on its assets, or repays one.",
     "steps": [
         {"t": "Board approval", "d": "Approve the borrowing and the security. Check whether members' consent is needed for the limit.", "f": [], "s": ["179", "180"]},
         {"t": "Execute the documents", "d": "Sign the loan and security documents with the lender.", "f": [], "s": ["77"]},
         {"t": "Register the charge", "d": "File the particulars of the charge within the time the Act gives. Late filing is possible on extra fees within a longer period.", "f": ["CHG-1"], "s": ["77", "78"]},
         {"t": "Keep the register of charges", "d": "Enter the charge in the company's register.", "f": [], "s": ["85"]},
         {"t": "On repayment, file satisfaction", "d": "File the satisfaction of the charge once the loan is repaid.", "f": ["CHG-4"], "s": ["82"]}],
     "check": ["Board resolution", "Members' resolution where required", "Charge documents signed", "CHG-1 filed in time", "Register of charges updated", "CHG-4 filed on repayment"],
     "drafts": ["creation of charge", "mortgage", "borrowing", "hypothecation"],
     "lodr": []},
    {"id": "annual-filing", "title": "Yearly filings after the AGM",
     "summary": "The returns a company files every year with the Registrar.",
     "when": "Each financial year, after the AGM.",
     "steps": [
         {"t": "File the financial statements", "d": "File with the Registrar after the AGM, within the time the Act gives.", "f": ["AOC-4"], "s": ["137"]},
         {"t": "File the annual return", "d": "File the annual return of members, directors and meetings.", "f": ["MGT-7"], "s": ["92"]},
         {"t": "File the auditor's appointment if due", "d": "If a new term begins, file the notice.", "f": ["ADT-1"], "s": ["139"]},
         {"t": "Return of deposits and exempt receipts", "d": "File the return showing money received that is not a deposit, by its due date.", "f": ["DPT-3"], "s": ["73"]},
         {"t": "Director KYC", "d": "Every director with a DIN files the yearly KYC by the due date.", "f": ["DIR-3"], "s": ["153"]},
         {"t": "Half-yearly return of dues to small enterprises", "d": "Companies that owe such dues beyond the stated period file the half-yearly return.", "f": ["MSME-1"], "s": ["405"]}],
     "check": ["AOC-4 filed", "MGT-7 filed", "DPT-3 filed", "Director KYC done", "MSME-1 half-yearly filed if applicable", "Register of members, directors and charges updated"],
     "drafts": ["annual return", "financial statements", "aoc", "mgt"],
     "lodr": []},
    {"id": "change-name", "title": "Changing the company's name",
     "summary": "Name approval, the resolution and the Registrar's fresh certificate.",
     "when": "The company wants a new name.",
     "steps": [
         {"t": "Reserve the new name", "d": "Apply for the new name with the Registrar.", "f": [], "s": ["4", "13"]},
         {"t": "Board and members' approval", "d": "Pass a board resolution and a special resolution, and alter the Memorandum and Articles.", "f": ["MGT-14"], "s": ["13", "14"]},
         {"t": "File the application", "d": "File the application for change of name with the Registrar.", "f": ["INC-24"], "s": ["13"]},
         {"t": "Receive the fresh certificate", "d": "The Registrar issues a fresh certificate. Update PAN, GST, bank, licences and letterheads.", "f": [], "s": ["13"]}],
     "check": ["Name reserved", "Special resolution passed", "MGT-14 filed", "INC-24 filed", "Fresh certificate received", "Authorities informed"],
     "drafts": ["change of name", "name change"],
     "lodr": [{"id": "30", "why": "Listed companies inform the exchanges."}]},
    {"id": "alter-moa-aoa", "title": "Altering the Memorandum or Articles",
     "summary": "A special resolution, a filing, and some changes that need extra approval.",
     "when": "The company changes its objects, capital clause or its Articles.",
     "steps": [
         {"t": "Board resolution", "d": "Approve the alteration and the notice for the members' meeting.", "f": [], "s": ["13", "14", "173"]},
         {"t": "Special resolution", "d": "Pass it at a general meeting. A change of objects may need a postal ballot.", "f": ["MGT-14"], "s": ["13", "14", "110"]},
         {"t": "File the altered documents", "d": "File the resolution and the altered documents with the Registrar within the time the Act gives.", "f": ["MGT-14"], "s": ["14", "117"]},
         {"t": "Registrar's confirmation where needed", "d": "Where the Act needs the Registrar or the Regional Director to confirm, wait for the order before acting on the change.", "f": ["INC-28"], "s": ["13"]}],
     "check": ["Board approval", "Special resolution passed", "Altered documents prepared", "MGT-14 filed", "Order received where required"],
     "drafts": ["alteration of memorandum", "alteration of articles", "amendment of articles"],
     "lodr": []},
    {"id": "authorised-capital", "title": "Increasing authorised capital",
     "summary": "Altering the capital clause and filing the notice.",
     "when": "The company wants authorised capital above its present limit.",
     "steps": [
         {"t": "Check the Articles", "d": "Confirm the Articles authorise an increase. If they do not, alter them first.", "f": [], "s": ["61", "14"]},
         {"t": "Board resolution", "d": "Approve the increase and call the general meeting.", "f": [], "s": ["61", "173"]},
         {"t": "Members' resolution", "d": "Pass the resolution at the meeting.", "f": ["MGT-14"], "s": ["61", "117"]},
         {"t": "File the notice and pay stamp duty", "d": "File the notice of the increase and the altered Memorandum.", "f": ["SH-7"], "s": ["64"]}],
     "check": ["Articles checked", "Board and members' resolutions", "SH-7 filed in time", "Stamp duty and fees paid", "Altered Memorandum filed"],
     "drafts": ["authorised capital", "authorized capital", "increase of capital"],
     "lodr": []},
    {"id": "related-party", "title": "Related party transactions",
     "summary": "Approval of the audit committee, the board and, where needed, the members.",
     "when": "The company proposes a transaction with a related party.",
     "steps": [
         {"t": "Identify the related party and the transaction", "d": "Check the list of related parties against the transaction.", "f": [], "s": ["2", "188"]},
         {"t": "Audit committee and board approval", "d": "Obtain approval, with interested directors keeping out of the vote.", "f": ["MBP-1"], "s": ["177", "184", "188"]},
         {"t": "Members' approval if the limits are crossed", "d": "Pass the resolution by the members who are not related.", "f": ["MGT-14"], "s": ["188"]},
         {"t": "Disclose", "d": "Record the transaction in the board's report and the register of contracts.", "f": ["MBP-4"], "s": ["189", "134"]}],
     "check": ["Related party list updated", "Arm's length and ordinary course tested", "Audit committee approval", "Board approval", "Members' approval if needed", "Register of contracts updated", "Disclosure in the board's report"],
     "drafts": ["related party"],
     "lodr": [{"id": "23", "why": "Related party transactions of listed companies."}]},
    {"id": "strike-off", "title": "Striking off a company's name",
     "summary": "Applying to the Registrar to remove the name of a company that has no business.",
     "when": "The company has not carried on business for the period the Act gives and has no liabilities.",
     "steps": [
         {"t": "Check that the company can apply", "d": "Confirm the conditions and prepare the statement of accounts.", "f": [], "s": ["248"]},
         {"t": "Pass resolutions", "d": "Pass a special resolution or obtain the consent of the required part of members by shares.", "f": ["MGT-14"], "s": ["248"]},
         {"t": "Prepare the documents", "d": "Affidavits by the directors, the indemnity bond, and the statement of accounts.", "f": [], "s": ["248"]},
         {"t": "File the application", "d": "File the application with the Registrar.", "f": ["STK-2"], "s": ["248"]}],
     "check": ["No business for the required period", "No assets or liabilities", "Special resolution passed", "Affidavits and indemnity bond signed", "STK-2 filed"],
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


def main():
    procs = list(PROCESSES)
    extra = Path("guide") / "processes.json"
    if extra.exists():
        try:
            procs += json.loads(extra.read_text(encoding="utf-8"))
        except ValueError as e:
            print(f"guide/processes.json could not be read: {e}")
    drafts = drafts_list()
    out = []
    for p in procs:
        hit = []
        for d in drafts:
            hay = " ".join(words(d["title"] + " " + d.get("category", "")))
            if any(k.lower() in hay for k in p.get("drafts", [])):
                hit.append({"id": d["id"], "title": d["title"], "category": d.get("category", "")})
        p = dict(p)
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
