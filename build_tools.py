#!/usr/bin/env python3
"""Builds two tool pages for the site:
  site/generators/index.html : document generators (fill in a form, download an editable Word file; all in the browser)
  site/excel/                : Excel utilities (registers, trackers and planners as .xlsx files) and their listing page
Needs the openpyxl library (installed by the workflow). Nothing here reads the Act; it only writes tools."""
import sys
from datetime import date
from pathlib import Path

SITE = Path("site")

CSS = r""":root{--paper:#e9eef5;--page:#ffffff;--ink:#101a2b;--muted:#4b5a70;--rule:#c9d3e1;--tape:#1c3d6e;--link:#1f4fb5;--mark:#ffe08a;--bar:#0f2747;--barink:#eef3fa;--gold:#c8a13a;
--serif:"Source Serif 4",Georgia,"Times New Roman",serif;--sans:"Zilla Slab",Rockwell,"Roboto Slab",Georgia,serif}
:root[data-theme=sepia]{--paper:#e6dcc3;--page:#f8f2e0;--ink:#2a2418;--muted:#625740;--rule:#d0c4a2;--tape:#7a4b12;--link:#7a3b0c;--bar:#3a2a14;--barink:#f6eedb}
:root[data-theme=dark]{--paper:#0b1220;--page:#121b2e;--ink:#e8eef8;--muted:#9aabc4;--rule:#26344f;--tape:#e0b84f;--link:#8fb4ff;--bar:#070d19;--barink:#e8eef8;--gold:#e0b84f;color-scheme:dark}
@media (prefers-color-scheme:dark){:root[data-theme=auto]{--paper:#0b1220;--page:#121b2e;--ink:#e8eef8;--muted:#9aabc4;--rule:#26344f;--tape:#e0b84f;--link:#8fb4ff;--bar:#070d19;--barink:#e8eef8;--gold:#e0b84f;color-scheme:dark}}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:var(--ink);font:400 1.05rem/1.55 var(--serif)}
a{color:var(--link)}
[hidden]{display:none!important}
:focus-visible{outline:2px solid var(--link);outline-offset:2px}
.bar{background:var(--bar);color:var(--barink);border-bottom:3px solid var(--gold)}
.bar-in{max-width:76rem;margin:auto;padding:.7rem 1.5rem;display:flex;align-items:center;gap:1.2rem;flex-wrap:wrap}
.brand{font:700 1.2rem var(--sans);color:var(--barink);text-decoration:none}
.bar nav{display:flex;gap:1rem;flex-wrap:wrap}
.bar nav a{color:var(--barink);text-decoration:none;font:600 .95rem var(--sans);opacity:.85;padding:.2rem 0;border-bottom:2px solid transparent}
.bar nav a:hover,.bar nav a[aria-current=page]{opacity:1;border-bottom-color:var(--gold)}
.themes{margin-left:auto;display:flex;gap:.4rem}
.sw{width:1.15rem;height:1.15rem;border-radius:50%;border:1px solid var(--muted);cursor:pointer;padding:0;background:var(--c)}
.sw[aria-pressed=true]{outline:2px solid var(--gold);outline-offset:2px}
.sw[data-t=light]{--c:#fff}.sw[data-t=sepia]{--c:#f3ecd7}.sw[data-t=dark]{--c:#0b1220}
main{max-width:76rem;margin:auto;padding:1.6rem 1.5rem 4rem}
h1{font:700 clamp(1.8rem,4vw,2.5rem)/1.1 var(--sans);color:var(--tape);margin:.2rem 0 .5rem}
:root[data-theme=dark] h1{color:var(--ink)}
h2{font:700 1.15rem var(--sans);margin:1.8rem 0 .7rem;padding-bottom:.35rem;border-bottom:1px solid var(--rule);color:var(--muted)}
.lede{max-width:46rem;color:var(--muted)}
.small{font-size:.92rem;color:var(--muted)}
.basis{font:500 .92rem var(--sans);color:var(--muted);margin:.2rem 0 1rem}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(17rem,1fr));gap:1rem}
.card{display:flex;flex-direction:column;gap:.35rem;background:var(--page);border:1px solid var(--rule);border-left:4px solid var(--tape);border-radius:3px;padding:1rem 1.1rem;text-decoration:none;color:var(--ink);box-shadow:0 12px 24px -20px rgba(15,39,71,.6)}
.card:hover{border-color:var(--tape)}
.card b{font:700 1.08rem var(--sans)}
.card span{color:var(--muted);font-size:.97rem}
.card i{font:normal 500 .82rem var(--sans);color:var(--muted);margin-top:auto;padding-top:.4rem}
.cols{display:grid;grid-template-columns:minmax(18rem,26rem) minmax(0,1fr);gap:2rem;align-items:start}
.form{background:var(--page);border:1px solid var(--rule);border-radius:3px;padding:.2rem 1.1rem 1.2rem;position:sticky;top:1rem;max-height:calc(100vh - 2rem);overflow:auto}
.fld{margin:.7rem 0}
.fld label{display:block;font:600 .9rem var(--sans);margin-bottom:.2rem}
.fld input,.fld textarea,.fld select{width:100%;font:inherit;font-size:1rem;padding:.5rem .6rem;border:1px solid var(--rule);border-radius:3px;background:var(--page);color:var(--ink)}
.fld input:focus,.fld textarea:focus,.fld select:focus{border-color:var(--link)}
.actions{display:flex;gap:.6rem;flex-wrap:wrap;align-items:center;margin-bottom:.8rem}
.btn{font:600 .95rem var(--sans);padding:.5rem 1rem;border:1px solid var(--tape);background:var(--page);color:var(--tape);border-radius:3px;cursor:pointer;text-decoration:none;display:inline-block}
.btn.pri{background:var(--tape);color:#fff}
:root[data-theme=dark] .btn.pri{color:#101a2b}
.msg{font:500 .9rem var(--sans);color:var(--muted)}
.paper{background:#fff;color:#111;border:1px solid var(--rule);box-shadow:0 18px 40px -28px rgba(15,39,71,.55);padding:2.4rem 2.6rem;font:400 1rem/1.6 "Times New Roman",Times,serif;min-height:20rem}
.paper p{margin:0 0 .7rem}
.paper .t{text-align:center;font-weight:700;text-decoration:underline;margin:1rem 0}
.paper .h,.paper .c{text-align:center}.paper .p{text-align:justify}.paper .r{text-align:right}.paper .sp{height:.8rem}
.check{margin-top:1.4rem;background:var(--page);border:1px solid var(--rule);border-left:4px solid var(--gold);padding:.2rem 1.2rem 1rem}
.check h2{border:0;margin:.9rem 0 .3rem}
.xl{display:grid;grid-template-columns:repeat(auto-fill,minmax(20rem,1fr));gap:1.1rem}
.xl .card{border-left-color:#1d7a4a}
.xl .card .btn{align-self:flex-start;margin-top:.6rem;background:#1d7a4a;border-color:#1d7a4a;color:#fff}
.xl .card ul{margin:.3rem 0 0;padding-left:1.1rem;color:var(--muted);font-size:.93rem}
footer{max-width:76rem;margin:0 auto;padding:1rem 1.5rem 3rem;color:var(--muted);font-size:.9rem;border-top:1px solid var(--rule)}
@media (max-width:860px){.cols{grid-template-columns:1fr}.form{position:static;max-height:none}.paper{padding:1.2rem}}
@media print{.bar,.form,.actions,.check,footer,.small,.lede,.basis,h1,#back{display:none!important}body{background:#fff}main{padding:0}.cols{display:block}.paper{border:0;box-shadow:none;padding:0}}

.mtabs{display:flex;gap:.5rem;flex-wrap:wrap;margin:.6rem 0 1rem}
.rlist{display:grid;gap:.5rem}
.rrow{display:flex;flex-direction:column;background:var(--page);border:1px solid var(--rule);border-left:4px solid var(--gold);border-radius:3px;padding:.6rem .9rem;text-decoration:none;color:var(--ink)}
.rrow:hover{border-color:var(--tape)}.rrow b{font:700 1rem var(--sans)}.rrow span{color:var(--muted);font-size:.9rem}
"""
GEN_JS = r""""use strict";
/* ---------- small helpers ---------- */
const $ = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const store = { get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }, set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} } };
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function fdate(s) { if (!s) return ""; const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s); if (!m) return s; return `${+m[3]} ${MONTHS[+m[2] - 1]} ${m[1]}`; }
function inrFmt(n) { n = String(Math.round(Number(n))); if (n.length <= 3) return n; const l = n.slice(-3); let r = n.slice(0, -3); r = r.replace(/\B(?=(\d{2})+(?!\d))/g, ","); return r + "," + l; }
const ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
const TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
function below1000(n) { let s = ""; if (n >= 100) { s += ONES[Math.floor(n / 100)] + " Hundred"; n %= 100; if (n) s += " and "; } if (n >= 20) { s += TENS[Math.floor(n / 10)]; if (n % 10) s += " " + ONES[n % 10]; } else if (n) s += ONES[n]; return s; }
function words(n) {
  n = Math.round(Number(n)); if (!isFinite(n) || n < 0) return ""; if (n === 0) return "Zero";
  const parts = []; const cr = Math.floor(n / 1e7); n %= 1e7; const lk = Math.floor(n / 1e5); n %= 1e5; const th = Math.floor(n / 1e3); n %= 1e3;
  if (cr) parts.push(words(cr) + " Crore"); if (lk) parts.push(below1000(lk) + " Lakh"); if (th) parts.push(below1000(th) + " Thousand"); if (n) parts.push(below1000(n));
  return parts.join(" ");
}

/* ---------- common company details (saved in this browser) ---------- */
const COMMON = [
  { id: "company", label: "Company name", ph: "ABC Private Limited", req: 1 },
  { id: "cin", label: "CIN", ph: "U12345MH2020PTC123456" },
  { id: "ro", label: "Registered office address", type: "textarea", ph: "Full address of the registered office" },
];

/* ---------- document blocks: T title, H heading, P justified, L left, C centred, R right, I indented, S space, PB page break ---------- */
const T = t => ({ k: "T", t }), H = t => ({ k: "H", t }), P = t => ({ k: "P", t }), L = t => ({ k: "L", t }), C = t => ({ k: "C", t }), R = t => ({ k: "R", t }), I = (t, n) => ({ k: "I", t, n: n || 1 }), S = () => ({ k: "S" }), PB = () => ({ k: "PB" });
const RES = t => P("**RESOLVED THAT** " + t);
const RESF = t => P("**RESOLVED FURTHER THAT** " + t);

function va(v, id, ph) { const x = (v[id] == null ? "" : String(v[id])).trim(); return x ? x : "[" + (ph || id) + "]"; }
function vd(v, id, ph) { const x = (v[id] || "").trim(); return x ? fdate(x) : "[" + (ph || "date") + "]"; }
function vmoney(v, id, ph) { const x = String(v[id] || "").replace(/[^0-9.]/g, ""); if (!x || isNaN(+x)) return "[" + (ph || "amount") + "]"; return `Rs. ${inrFmt(x)}/- (Rupees ${words(x)} only)`; }
function vnum(v, id, ph) { const x = String(v[id] || "").replace(/[^0-9.]/g, ""); if (!x || isNaN(+x)) return "[" + (ph || "number") + "]"; return `${inrFmt(x)} (${words(x)})`; }
function header(v) {
  const o = [C("**" + va(v, "company", "Company name").toUpperCase() + "**")];
  if ((v.cin || "").trim()) o.push(C("CIN: " + v.cin.trim()));
  if ((v.ro || "").trim()) o.push(C("Registered Office: " + v.ro.trim().replace(/\n+/g, ", ")));
  o.push(S()); return o;
}
function signoff(v, who, des, extra) {
  return [S(), R("For **" + va(v, "company", "Company name") + "**"), S(), S(), R("**" + who + "**"), R(des)].concat(extra || []);
}
const PLACE = v => (v.place || "").trim() || "the registered office of the Company";
const PURPOSES = [
  ["director", "Appoint an additional director"], ["office", "Shift registered office within the same city"], ["bank", "Open a bank account and name signatories"],
  ["efile", "Authorise a person to sign and file e-forms"], ["allot", "Allot shares"], ["loan", "Avail a loan or credit facility"],
  ["auditor", "Appoint the first auditor"], ["agm", "Convene the Annual General Meeting"], ["fs", "Approve financial statements and Board's report"],
];
const showFor = (...p) => v => p.indexOf(v.purpose) >= 0;

/* ---------- the generators ---------- */
const GENS = [
  {
    id: "board", group: "Board resolutions", title: "Board resolution", short: "Certified true copy of a Board resolution for nine common purposes.",
    basis: "Companies Act, 2013, section 179 and the section named in each purpose",
    fields: [
      { id: "purpose", label: "Purpose", type: "select", opts: PURPOSES },
      { id: "mdate", label: "Date of Board meeting", type: "date" }, { id: "mtime", label: "Time of meeting", ph: "11:00 a.m." }, { id: "place", label: "Place of meeting", ph: "Registered office" },
      { id: "name", label: "Name of the person", ph: "Full name", show: showFor("director", "efile") }, { id: "din", label: "DIN", ph: "00000000", show: showFor("director") },
      { id: "desig", label: "Designation", ph: "Director / Company Secretary", show: showFor("efile") },
      { id: "effdate", label: "Effective date", type: "date", show: showFor("director", "office") },
      { id: "oldro", label: "Present registered office", type: "textarea", show: showFor("office") }, { id: "newro", label: "New registered office", type: "textarea", show: showFor("office") }, { id: "city", label: "City or town (same for both)", show: showFor("office") },
      { id: "bank", label: "Name of bank", show: showFor("bank") }, { id: "branch", label: "Branch", show: showFor("bank") }, { id: "signs", label: "Authorised signatories (one per line)", type: "textarea", show: showFor("bank") },
      { id: "mode", label: "Mode of operation", type: "select", opts: [["singly", "Singly (either or anyone)"], ["jointly", "Jointly (any two)"]], show: showFor("bank") },
      { id: "shares", label: "Number of shares", type: "number", show: showFor("allot") }, { id: "class", label: "Class of shares", ph: "equity shares", def: "equity shares", show: showFor("allot") },
      { id: "fv", label: "Face value per share (Rs.)", type: "number", show: showFor("allot") }, { id: "price", label: "Issue price per share (Rs.)", type: "number", show: showFor("allot") },
      { id: "allottee", label: "Name of allottee(s)", type: "textarea", show: showFor("allot") }, { id: "apprdate", label: "Date of members' approval (if required)", type: "date", show: showFor("allot") },
      { id: "lender", label: "Name of lender", show: showFor("loan") }, { id: "amount", label: "Amount (Rs.)", type: "number", show: showFor("loan") }, { id: "facility", label: "Type of facility", ph: "term loan / cash credit", show: showFor("loan") },
      { id: "sanction", label: "Sanction letter date", type: "date", show: showFor("loan") }, { id: "signer", label: "Authorised signatory for loan documents", show: showFor("loan") },
      { id: "firm", label: "Auditor's firm name", show: showFor("auditor") }, { id: "frn", label: "Firm registration number", show: showFor("auditor") }, { id: "remun", label: "Remuneration (Rs.)", type: "number", show: showFor("auditor") },
      { id: "agmno", label: "Number of the AGM", ph: "5th", show: showFor("agm") }, { id: "agmdate", label: "Date of AGM", type: "date", show: showFor("agm") }, { id: "agmtime", label: "Time of AGM", ph: "11:00 a.m.", show: showFor("agm") }, { id: "venue", label: "Venue of AGM", type: "textarea", show: showFor("agm") },
      { id: "fyend", label: "Financial year ended on", type: "date", show: showFor("fs", "agm") }, { id: "signer1", label: "First director authorised to sign", show: showFor("fs") }, { id: "signer2", label: "Second director or officer authorised to sign", show: showFor("fs") },
      { id: "auth", label: "Person authorised to give effect (name and designation)", ph: "Mr. A, Director", req: 1 },
      { id: "sname", label: "Name of the signing director", req: 1 }, { id: "sdin", label: "DIN of the signing director" },
    ],
    notes: ["Check the Articles of Association for any extra approval the purpose needs.", "Allotment and borrowing may also need members' approval and later ROC filings (PAS-3, CHG-1). The resolution only authorises them.", "Attach the resolution to the minutes of the meeting and keep it in the minute book."],
    build(v) {
      const p = v.purpose, co = va(v, "company", "Company name"), auth = va(v, "auth", "authorised person");
      const out = header(v);
      out.push(T("CERTIFIED TRUE COPY OF THE RESOLUTION PASSED AT THE MEETING OF THE BOARD OF DIRECTORS OF " + co.toUpperCase() + " HELD ON " + vd(v, "mdate", "date of meeting").toUpperCase() + (v.mtime ? " AT " + v.mtime.toUpperCase() : "") + " AT " + PLACE(v).toUpperCase()));
      const give = () => RESF(`any director or the Company Secretary of the Company, or ${auth}, be and is hereby authorised to do all such acts, deeds and things, sign and file all forms and documents with the Registrar of Companies or any other authority, and to settle any question or difficulty that may arise, to give effect to this resolution.`);
      if (p === "director") {
        out.push(P("**1. Appointment of an Additional Director**"));
        out.push(RES(`pursuant to Section 161(1) of the Companies Act, 2013 and the Articles of Association of the Company, ${va(v, "name", "name")} (DIN: ${va(v, "din", "DIN")}), who has given his or her consent to act as a director (Form DIR-2) and a declaration that he or she is not disqualified (Form DIR-8), be and is hereby appointed as an Additional Director of the Company with effect from ${vd(v, "effdate", "effective date")}, to hold office up to the date of the next Annual General Meeting or the last date on which that meeting should have been held, whichever is earlier.`));
        out.push(RESF(`${auth} be and is hereby authorised to file Form DIR-12 with the Registrar of Companies within thirty days of the appointment and to do all things needed to give effect to this resolution.`));
      } else if (p === "office") {
        out.push(P("**1. Shifting of the Registered Office within the same city**"));
        out.push(RES(`pursuant to Section 12 of the Companies Act, 2013 and the Articles of Association of the Company, the registered office of the Company be shifted from ${va(v, "oldro", "present address")} to ${va(v, "newro", "new address")}, both within the local limits of ${va(v, "city", "city or town")}, with effect from ${vd(v, "effdate", "effective date")}.`));
        out.push(RESF(`${auth} be and is hereby authorised to file the notice of change of registered office in Form INC-22 with the Registrar of Companies within the prescribed time, with the required documents, and to inform all concerned authorities and stakeholders.`));
      } else if (p === "bank") {
        const names = (v.signs || "").split(/\n+/).map(s => s.trim()).filter(Boolean);
        out.push(P("**1. Opening of a Bank Account**"));
        out.push(RES(`a current account of the Company be opened with ${va(v, "bank", "bank")}, ${va(v, "branch", "branch")} Branch, and the bank be authorised to honour cheques, instructions and electronic payment orders drawn or given on the account in accordance with this resolution.`));
        out.push(RESF(`the following persons be and are hereby authorised to operate the account ${v.mode === "jointly" ? "jointly, any two of them acting together" : "singly, any one of them acting alone"}, including to sign cheques, give payment instructions and operate net banking:`));
        (names.length ? names : ["[name of signatory]"]).forEach((n, i) => out.push(I(`${i + 1}. ${n}`)));
        out.push(RESF("the bank's standard form of mandate, together with a copy of this resolution certified by a director, be given to the bank, and the bank be informed of any change in the authorised persons by a fresh resolution."));
      } else if (p === "efile") {
        out.push(P("**1. Authority to sign and file e-forms**"));
        out.push(RES(`${va(v, "name", "name")}, ${va(v, "desig", "designation")} of the Company, be and is hereby authorised to sign, certify and file, by affixing his or her digital signature, all e-forms, returns and documents required to be filed with the Registrar of Companies, the Ministry of Corporate Affairs and other regulatory authorities, and to represent the Company before them in connection with those filings.`));
        out.push(RESF("this authority shall continue until it is withdrawn or replaced by a further resolution of the Board."));
      } else if (p === "allot") {
        out.push(P("**1. Allotment of Shares**"));
        out.push(RES(`subject to the provisions of the Companies Act, 2013 and the rules made under it, the Memorandum and Articles of Association of the Company and the approval of the members${v.apprdate ? " given on " + fdate(v.apprdate) : ""} (where required), ${vnum(v, "shares", "number of shares")} ${va(v, "class", "class")} of Rs. ${va(v, "fv", "face value")} each be and are hereby allotted at an issue price of Rs. ${va(v, "price", "price")} per share to ${va(v, "allottee", "name of allottee")}, against receipt of the full application money in the Company's bank account.`));
        out.push(RESF(`the share certificate(s) in Form SH-1 be issued within two months of the date of allotment, the Register of Members be updated, and the return of allotment in Form PAS-3 be filed with the Registrar of Companies within the time prescribed (fifteen days for a private placement, and thirty days in other cases).`));
      } else if (p === "loan") {
        out.push(P("**1. Credit facility from " + va(v, "lender", "lender") + "**"));
        out.push(RES(`the Company do avail ${va(v, "facility", "type of facility")} of ${vmoney(v, "amount", "amount")} from ${va(v, "lender", "lender")}, on the terms and conditions set out in the sanction letter dated ${vd(v, "sanction", "sanction date")}, and the same be accepted.`));
        out.push(RESF(`${va(v, "signer", "authorised signatory")} be and is hereby authorised to sign and execute the loan agreement, security documents and all related papers on behalf of the Company, to create the security required, and to file the particulars of any charge created with the Registrar of Companies within thirty days of its creation.`));
        out.push(P("It was noted that the total borrowings of the Company, including this facility, are within the limit approved by the members under Section 180(1)(c) of the Companies Act, 2013, where that section applies."));
      } else if (p === "auditor") {
        out.push(P("**1. Appointment of the First Auditor**"));
        out.push(RES(`pursuant to Section 139(6) of the Companies Act, 2013, ${va(v, "firm", "name of firm")}, Chartered Accountants (Firm Registration No. ${va(v, "frn", "FRN")}), who have given their written consent and a certificate that they are eligible under Section 141, be and are hereby appointed as the first auditor of the Company, to hold office from the date of this resolution until the conclusion of the first Annual General Meeting, at a remuneration of ${vmoney(v, "remun", "remuneration")} plus applicable taxes and reimbursement of out-of-pocket expenses.`));
        out.push(RESF(`${auth} be and is hereby authorised to intimate the auditor and to file Form ADT-1 with the Registrar of Companies within the prescribed time.`));
      } else if (p === "agm") {
        out.push(P("**1. Convening the Annual General Meeting**"));
        out.push(RES(`the ${va(v, "agmno", "number")} Annual General Meeting of the Company for the financial year ended ${vd(v, "fyend", "year end")} be held on ${vd(v, "agmdate", "AGM date")} at ${va(v, "agmtime", "time")} at ${va(v, "venue", "venue")}, and the notice of the meeting, together with the explanatory statement, as placed before the Board, be and is hereby approved.`));
        out.push(RESF("the notice, the Annual Report and other documents be sent to the members, directors and the auditors, and to every other person entitled to receive them, at least twenty-one clear days before the date of the meeting, unless the meeting is called at shorter notice with the consent required by Section 101."));
      } else if (p === "fs") {
        out.push(P("**1. Approval of Financial Statements and the Board's Report**"));
        out.push(RES(`the audited financial statements of the Company for the financial year ended ${vd(v, "fyend", "year end")}, consisting of the Balance Sheet, the Statement of Profit and Loss, the Cash Flow Statement and the notes, together with the Auditor's Report, be and are hereby approved.`));
        out.push(RESF(`the Board's Report, with its annexures, prepared under Section 134 of the Companies Act, 2013, be and is hereby approved, and ${va(v, "signer1", "first director")} and ${va(v, "signer2", "second signatory")} be and are hereby authorised to sign the financial statements and the Board's Report for and on behalf of the Board.`));
      }
      if (p !== "efile") out.push(give());
      out.push(S(), L("Certified to be a true copy."));
      return out.concat(signoff(v, va(v, "sname", "name of director"), "Director" + ((v.sdin || "").trim() ? " (DIN: " + v.sdin.trim() + ")" : ""), [L("Place: " + (v.place || "") ), L("Date: " + fdate(v.mdate))]));
    },
  },
  {
    id: "dividend", group: "Board resolutions", title: "Interim dividend resolution", short: "Board resolution declaring an interim dividend.",
    basis: "Section 123(3) to (5) and Section 127, Companies Act, 2013; Companies (Declaration and Payment of Dividend) Rules, 2014",
    fields: [
      { id: "mdate", label: "Date of Board meeting", type: "date" }, { id: "fy", label: "Financial year", ph: "2026-27" },
      { id: "rate", label: "Dividend per share (Rs.)", type: "number" }, { id: "fv", label: "Face value per share (Rs.)", type: "number" }, { id: "total", label: "Total dividend (Rs.)", type: "number" },
      { id: "rdate", label: "Record date", type: "date" }, { id: "pdate", label: "Payment on or before", type: "date" }, { id: "bank", label: "Bank for the separate dividend account" },
      { id: "auth", label: "Person authorised (name and designation)" }, { id: "sname", label: "Name of the signing director" },
    ],
    notes: ["The interim dividend may be paid only out of the surplus in the profit and loss account and the profits of the current year. If there is a loss, the rate cannot be higher than the average of the last three years' dividend.", "The amount must be moved to a separate bank account within five days and paid to members within thirty days of declaration (Section 127).", "Tax deduction and the details of the shares eligible are not covered here."],
    build(v) {
      const out = header(v);
      out.push(T("RESOLUTION PASSED AT THE MEETING OF THE BOARD OF DIRECTORS HELD ON " + vd(v, "mdate", "date").toUpperCase()));
      out.push(P("**Declaration of Interim Dividend for the financial year " + va(v, "fy", "year") + "**"));
      out.push(RES(`pursuant to Section 123(3) of the Companies Act, 2013, the Companies (Declaration and Payment of Dividend) Rules, 2014 and the Articles of Association of the Company, an interim dividend of Rs. ${va(v, "rate", "rate")} per equity share of the face value of Rs. ${va(v, "fv", "face value")} each, amounting in all to ${vmoney(v, "total", "total amount")}, be and is hereby declared for the financial year ${va(v, "fy", "year")}, out of the surplus in the Statement of Profit and Loss and the profits of the current financial year.`));
      out.push(RESF(`the dividend be paid to those members whose names appear in the Register of Members as on ${vd(v, "rdate", "record date")} (the record date), on or before ${vd(v, "pdate", "payment date")}, and in any case within thirty days of the date of declaration.`));
      out.push(RESF(`the total amount of the dividend be deposited in a separate bank account with ${va(v, "bank", "bank")}, a scheduled bank, within five days from the date of declaration, and be used only for paying the dividend.`));
      out.push(RESF(`${va(v, "auth", "authorised person")} be and is hereby authorised to take all steps needed to give effect to this resolution, including deduction of tax at source as required by law, the issue of dividend warrants or electronic payment, and the transfer of any unpaid amount to the Unpaid Dividend Account.`));
      return out.concat(signoff(v, va(v, "sname", "name"), "Director"));
    },
  },
  {
    id: "remun", group: "Board resolutions", title: "Director remuneration resolution", short: "Board resolution approving the remuneration of a managing, whole-time or other director.",
    basis: "Sections 196, 197 and 198 and Schedule V, Companies Act, 2013",
    fields: [
      { id: "mdate", label: "Date of Board meeting", type: "date" }, { id: "name", label: "Name of the director" }, { id: "din", label: "DIN" },
      { id: "desig", label: "Position", type: "select", opts: [["Managing Director", "Managing Director"], ["Whole-time Director", "Whole-time Director"], ["Director", "Director (other)"]] },
      { id: "from", label: "Term starts on", type: "date" }, { id: "years", label: "Term in years", type: "number" },
      { id: "salary", label: "Monthly salary (Rs.)", type: "number" }, { id: "perq", label: "Perquisites and other terms", type: "textarea", ph: "Housing, car, medical cover, bonus as per policy" },
      { id: "nrc", label: "Recommended by the Nomination and Remuneration Committee?", type: "select", opts: [["yes", "Yes"], ["no", "Not applicable"]] },
      { id: "gmdate", label: "General meeting for members' approval (if known)", type: "date" }, { id: "sname", label: "Name of the signing director" },
    ],
    notes: ["Sections 196 and 197 and Schedule V apply to public companies. A private company has no cap, but its Articles may need the members' approval.", "Remuneration above the Section 197 limits, or in the case of inadequate profits, needs the members' approval, and sometimes the Central Government's.", "File Form MGT-14 for a special resolution, and Form DIR-12 or MR-1 as applicable."],
    build(v) {
      const out = header(v);
      out.push(T("RESOLUTION PASSED AT THE MEETING OF THE BOARD OF DIRECTORS HELD ON " + vd(v, "mdate", "date").toUpperCase()));
      out.push(P("**Remuneration of " + va(v, "desig", "position") + "**"));
      if (v.nrc === "yes") out.push(P("The Chairman informed the Board that the Nomination and Remuneration Committee has reviewed and recommended the terms below."));
      out.push(RES(`pursuant to Sections 196, 197, 198 and other applicable provisions of the Companies Act, 2013, read with Schedule V, and the Articles of Association of the Company, and subject to the approval of the members in general meeting${v.gmdate ? " to be held on " + fdate(v.gmdate) : ""}, ${va(v, "name", "name")} (DIN: ${va(v, "din", "DIN")}) be and is hereby appointed or continued as ${va(v, "desig", "position")} of the Company for a term of ${va(v, "years", "years")} years from ${vd(v, "from", "start date")}, on a monthly salary of ${vmoney(v, "salary", "salary")}, together with the following perquisites and terms: ${va(v, "perq", "perquisites")}.`));
      out.push(RESF("the above remuneration be paid as minimum remuneration to the director in any financial year in which the Company has no profits or inadequate profits, within the limits of Schedule V to the Act."));
      out.push(RESF("the Board be authorised to alter and vary the terms within the limits permitted by the Act, and any director or the Company Secretary be authorised to file the forms required and to do all things needed to give effect to this resolution."));
      return out.concat(signoff(v, va(v, "sname", "name"), "Director"));
    },
  },
  {
    id: "loan", group: "Board resolutions", title: "Loan, guarantee or investment resolution", short: "Board resolution under Section 185 or Section 186.",
    basis: "Sections 185 and 186, Companies Act, 2013; Companies (Meetings of Board and its Powers) Rules, 2014",
    fields: [
      { id: "mdate", label: "Date of Board meeting", type: "date" },
      { id: "kind", label: "Type of transaction", type: "select", opts: [["loan", "Loan"], ["guarantee", "Guarantee"], ["security", "Providing security"], ["investment", "Acquire securities (investment)"]] },
      { id: "sect", label: "Section", type: "select", opts: [["186", "Section 186: any other body corporate or person"], ["185", "Section 185: a director or a person in whom a director is interested"]] },
      { id: "party", label: "Name of the borrower or investee" }, { id: "amount", label: "Amount (Rs.)", type: "number" }, { id: "rate", label: "Rate of interest per year (%)", type: "number" },
      { id: "tenure", label: "Tenure", ph: "3 years" }, { id: "purpose", label: "Purpose", type: "textarea", ph: "To meet working-capital needs" },
      { id: "gmdate", label: "Date of members' special resolution (if passed)", type: "date" }, { id: "sname", label: "Name of the signing director" },
    ],
    notes: ["Section 185: a loan, guarantee or security to a director or a person in whom the director is interested needs a special resolution of the members and must be used for the borrower's principal business.", "Section 186: the Board's approval must have the consent of all directors present. A special resolution is needed if the total goes beyond the limits in Section 186(2).", "The interest rate must not be lower than the prevailing yield of a Government security of a similar tenor. Record the transaction in the register (Form MBP-2).", "Some transactions are exempt (for example, a wholly owned subsidiary). Check Section 186(11) before using this draft."],
    build(v) {
      const out = header(v); const k = { loan: "give a loan", guarantee: "give a guarantee", security: "provide security", investment: "acquire securities" }[v.kind] || "give a loan"; const s185 = v.sect === "185";
      out.push(T("RESOLUTION PASSED AT THE MEETING OF THE BOARD OF DIRECTORS HELD ON " + vd(v, "mdate", "date").toUpperCase()));
      out.push(P("**" + (s185 ? "Section 185" : "Section 186") + ": " + k.replace(/^\w/, c => c.toUpperCase()) + " to " + va(v, "party", "party") + "**"));
      if (s185) {
        out.push(RES(`pursuant to Section 185 of the Companies Act, 2013, and subject to the approval of the members by a special resolution${v.gmdate ? " passed on " + fdate(v.gmdate) : ""}, the Company be and is hereby authorised to ${k} of ${vmoney(v, "amount", "amount")} to ${va(v, "party", "party")}, in whom a director of the Company is interested, for a tenure of ${va(v, "tenure", "tenure")}, at an interest rate of ${va(v, "rate", "rate")}% per annum, which is not lower than the prevailing yield of the Government security closest to the tenor of the loan.`));
        out.push(RESF(`the amount be utilised by the borrower for its principal business activities only, being: ${va(v, "purpose", "purpose")}.`));
      } else {
        out.push(RES(`pursuant to Section 186 of the Companies Act, 2013 and the rules made under it, the Board, with the consent of all the directors present at the meeting, approves that the Company ${k} of ${vmoney(v, "amount", "amount")} to or in ${va(v, "party", "party")}, for ${va(v, "purpose", "purpose")}, for a tenure of ${va(v, "tenure", "tenure")}${v.kind === "loan" ? ", at an interest rate of " + va(v, "rate", "rate") + "% per annum, which is not lower than the prevailing yield of the Government security closest to the tenor of the loan" : ""}.`));
        out.push(RESF(`the Board notes that the aggregate of loans, guarantees, securities and investments, including this one, is within the limits of Section 186(2), or within the limits approved by the members by a special resolution${v.gmdate ? " passed on " + fdate(v.gmdate) : ""}.`));
      }
      out.push(RESF("any director or the Company Secretary be authorised to sign the agreements and documents, to make the entry in the register under Section 186(9) (Form MBP-2), and to do all things needed to give effect to this resolution."));
      return out.concat(signoff(v, va(v, "sname", "name"), "Director"));
    },
  },
  {
    id: "bmnotice", group: "Meetings", title: "Notice of Board meeting", short: "Notice with the agenda, for the directors.",
    basis: "Section 173(3), Companies Act, 2013; Secretarial Standard 1",
    fields: [
      { id: "no", label: "Meeting number", ph: "3rd" }, { id: "fy", label: "Financial year", ph: "2026-27" }, { id: "ndate", label: "Date of notice", type: "date" },
      { id: "mdate", label: "Date of meeting", type: "date" }, { id: "mtime", label: "Time", ph: "11:00 a.m." }, { id: "venue", label: "Venue", type: "textarea", ph: "Registered office, or via video conferencing" },
      { id: "agenda", label: "Agenda items (one per line)", type: "textarea", def: "To confirm the minutes of the previous meeting\nTo note the compliance report\nTo consider any other business with the permission of the Chairman" },
      { id: "sname", label: "Name of the person issuing the notice" }, { id: "sdes", label: "Designation", ph: "Company Secretary" },
    ],
    notes: ["The notice must be sent at least seven days before the meeting, in writing, by hand, post or email, to every director at the address he or she has given.", "A shorter notice is possible if at least one independent director is present, or, if there is none, the decision is ratified by one independent director later.", "Send the agenda and the notes on each item with the notice. A short notice for items that are not confidential still needs the notes on the agenda."],
    build(v) {
      const out = header(v); const items = (v.agenda || "").split(/\n+/).map(s => s.trim()).filter(Boolean);
      out.push(T("NOTICE OF THE " + va(v, "no", "number").toUpperCase() + " MEETING OF THE BOARD OF DIRECTORS FOR THE FINANCIAL YEAR " + va(v, "fy", "year")));
      out.push(R("Date: " + vd(v, "ndate", "date of notice")));
      out.push(L("To,"), L("All the Directors of the Company"), S());
      out.push(P(`**NOTICE** is hereby given under Section 173(3) of the Companies Act, 2013 and Secretarial Standard 1 that the ${va(v, "no", "number")} meeting of the Board of Directors of the Company for the financial year ${va(v, "fy", "year")} will be held on **${vd(v, "mdate", "meeting date")}** at **${va(v, "mtime", "time")}** at ${va(v, "venue", "venue")}, to transact the business set out in the agenda below.`));
      out.push(P("**Agenda**"));
      (items.length ? items : ["[agenda item]"]).forEach((a, i) => out.push(I(`${i + 1}. ${a}`)));
      out.push(S(), P("The notes on the agenda items are enclosed. Directors who wish to take part through video conferencing or other audio-visual means may inform the Company Secretary in advance. Please confirm your attendance."));
      return out.concat([S(), L("By order of the Board"), R("For **" + va(v, "company", "Company name") + "**"), S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation"))]);
    },
  },
  {
    id: "agmnotice", group: "Meetings", title: "Notice of Annual General Meeting", short: "Notice for members, with ordinary and special business.",
    basis: "Sections 96, 101, 102 and 105, Companies Act, 2013; Secretarial Standard 2",
    fields: [
      { id: "no", label: "Number of the AGM", ph: "5th" }, { id: "ndate", label: "Date of notice", type: "date" }, { id: "mdate", label: "Date of AGM", type: "date" }, { id: "mtime", label: "Time", ph: "11:00 a.m." },
      { id: "venue", label: "Venue", type: "textarea" }, { id: "fyend", label: "Financial year ended on", type: "date" }, { id: "retire", label: "Director retiring by rotation and seeking reappointment" },
      { id: "special", label: "Special business (one item per line, then add 'Ordinary' or 'Special' after a | sign)", type: "textarea", ph: "To alter the Articles of Association | Special" },
      { id: "sname", label: "Name of the person issuing the notice" }, { id: "sdes", label: "Designation", ph: "Director / Company Secretary" },
    ],
    notes: ["Send the notice at least 21 clear days before the meeting to every member, director and auditor. A shorter notice needs the consent of members holding at least 95% of the voting power.", "The explanatory statement for each item of special business must be attached (Section 102). The placeholders are left for you to fill in.", "Add the e-voting details if the Company has to give them. This draft doesn't include them."],
    build(v) {
      const out = header(v); const sp = (v.special || "").split(/\n+/).map(s => s.trim()).filter(Boolean);
      out.push(T("NOTICE OF THE " + va(v, "no", "number").toUpperCase() + " ANNUAL GENERAL MEETING"));
      out.push(R("Date: " + vd(v, "ndate", "date of notice")));
      out.push(P(`**NOTICE** is hereby given that the ${va(v, "no", "number")} Annual General Meeting of the members of **${va(v, "company", "Company name")}** will be held on **${vd(v, "mdate", "meeting date")}** at **${va(v, "mtime", "time")}** at ${va(v, "venue", "venue")}, to transact the following business:`));
      out.push(P("**ORDINARY BUSINESS**"));
      let n = 1;
      out.push(P(`**${n++}.** To receive, consider and adopt the audited financial statements of the Company for the financial year ended ${vd(v, "fyend", "year end")}, together with the reports of the Board of Directors and the Auditors.`));
      if ((v.retire || "").trim()) out.push(P(`**${n++}.** To appoint a director in place of ${v.retire.trim()}, who retires by rotation under Section 152(6) of the Companies Act, 2013 and, being eligible, offers himself or herself for reappointment.`));
      if (sp.length) {
        out.push(P("**SPECIAL BUSINESS**"));
        sp.forEach(line => { const a = line.split("|"); const t = a[0].trim(); const k = /special/i.test(a[1] || "") ? "Special" : "Ordinary"; out.push(P(`**${n++}.** To consider and, if thought fit, to pass the following resolution as an ${k} Resolution: ${t}`)); });
      }
      out.push(S(), L("By order of the Board"), R("For **" + va(v, "company", "Company name") + "**"), S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")));
      out.push(S(), P("**Notes**"));
      out.push(P("1. A member entitled to attend and vote at the meeting may appoint a proxy to attend and vote instead of himself or herself, and the proxy need not be a member. The proxy form, duly completed and signed, must reach the registered office not less than 48 hours before the meeting. A person can act as proxy for not more than fifty members holding, in total, not more than ten per cent of the total share capital carrying voting rights."));
      out.push(P("2. Corporate members must send a certified copy of the Board resolution authorising their representative to attend and vote."));
      out.push(P("3. An explanatory statement under Section 102 for each item of special business is annexed to this notice."));
      if (sp.length) { out.push(S(), P("**EXPLANATORY STATEMENT UNDER SECTION 102 OF THE COMPANIES ACT, 2013**")); let m = (v.retire || "").trim() ? 3 : 2; sp.forEach(line => { m++; out.push(P(`**Item No. ${m}:** ${line.split("|")[0].trim()}`)); out.push(P("[Insert the reasons, the interest of the directors and key managerial personnel, and the other facts the members need to decide.]")); }); }
      return out;
    },
  },
  {
    id: "minutes", group: "Meetings", title: "Minutes of Board meeting", short: "A complete set of minutes with attendance, quorum and the decisions.",
    basis: "Section 118, Companies Act, 2013; Secretarial Standard 1",
    fields: [
      { id: "no", label: "Meeting number", ph: "3rd" }, { id: "mdate", label: "Date of meeting", type: "date" }, { id: "place", label: "Place", ph: "Registered office" }, { id: "start", label: "Started at", ph: "11:00 a.m." }, { id: "end", label: "Ended at", ph: "12:30 p.m." },
      { id: "chair", label: "Chairman of the meeting" }, { id: "present", label: "Directors present (one per line)", type: "textarea" }, { id: "attend", label: "In attendance (one per line)", type: "textarea" },
      { id: "leave", label: "Leave of absence granted to (one per line, if any)", type: "textarea" },
      { id: "items", label: "Items: title, then the decision after a | sign (one item per line)", type: "textarea", def: "Confirmation of the minutes of the previous meeting | The minutes of the meeting held on [date] were confirmed and signed by the Chairman.\nApproval of the compliance report | The Board noted the compliance report placed before it." },
      { id: "sdate", label: "Date of signing the minutes", type: "date" },
    ],
    notes: ["Minutes should be circulated in draft within 15 days and signed within 30 days of the meeting (Secretarial Standard 1).", "Record dissent or abstention of any director by name. Write the decisions as resolutions where they have to be filed.", "Enter the signed minutes in the minutes book, with each page initialled and the last page signed and dated."],
    build(v) {
      const list = id => (v[id] || "").split(/\n+/).map(s => s.trim()).filter(Boolean);
      const out = header(v); const pres = list("present"), att = list("attend"), lv = list("leave"), items = list("items");
      out.push(T(`MINUTES OF THE ${va(v, "no", "number").toUpperCase()} MEETING OF THE BOARD OF DIRECTORS HELD ON ${vd(v, "mdate", "date").toUpperCase()} AT ${va(v, "start", "time")} AT ${PLACE(v).toUpperCase()}`));
      out.push(P("**Directors present**"));
      (pres.length ? pres : ["[names of directors]"]).forEach((x, i) => out.push(I(`${i + 1}. ${x}`)));
      if (lv.length) { out.push(P("**Leave of absence was granted to**")); lv.forEach((x, i) => out.push(I(`${i + 1}. ${x}`))); }
      if (att.length) { out.push(P("**In attendance**")); att.forEach((x, i) => out.push(I(`${i + 1}. ${x}`))); }
      out.push(S(), P(`${va(v, "chair", "name")} took the chair and, the required quorum being present, called the meeting to order at ${va(v, "start", "time")}. The Chairman confirmed that the notice and the agenda had been sent to all the directors in time.`));
      (items.length ? items : ["[Item] | [Decision]"]).forEach((it, i) => { const a = it.split("|"); out.push(P(`**Item ${i + 1}: ${a[0].trim()}**`)); out.push(P((a[1] || "").trim() || "[Record the discussion and the decision.]")); });
      out.push(S(), P(`There being no other business, the meeting ended at ${va(v, "end", "time")} with a vote of thanks to the Chair.`));
      return out.concat([S(), S(), R("**" + va(v, "chair", "Chairman") + "**"), R("Chairman"), L("Place: " + (v.place || "")), L("Date of signing: " + fdate(v.sdate))]);
    },
  },
  {
    id: "authority", group: "Authorisation", title: "Letter of authority", short: "Authorises a person to act for the company for stated purposes and a stated period.",
    basis: "Section 179 and the Articles of Association; sections referred to in the purpose",
    fields: [
      { id: "date", label: "Date of letter", type: "date" }, { id: "to", label: "Addressed to", ph: "The Manager, ABC Bank, or 'Whom it may concern'" }, { id: "name", label: "Name of the authorised person" }, { id: "idno", label: "PAN or other ID of that person" },
      { id: "desig", label: "Designation or relationship" }, { id: "scope", label: "What the person is authorised to do", type: "textarea", ph: "Submit and collect documents; sign applications; deal with officials" },
      { id: "from", label: "Valid from", type: "date" }, { id: "upto", label: "Valid until", type: "date" }, { id: "res", label: "Board resolution date (if any)", type: "date" }, { id: "sname", label: "Name of the person signing for the company" }, { id: "sdes", label: "Designation", ph: "Director" },
    ],
    notes: ["Attach a certified copy of the Board resolution and an ID proof of the authorised person if the other side asks for it.", "Keep the scope narrow and the period short. Banks and registries often want the Board resolution itself."],
    build(v) {
      const out = header(v);
      out.push(R("Date: " + vd(v, "date", "date")));
      out.push(L("To,"), L(va(v, "to", "addressee")), S());
      out.push(P("**Subject: Letter of authority**"));
      out.push(P(`We, **${va(v, "company", "Company name")}**, hereby authorise **${va(v, "name", "name")}**, ${va(v, "desig", "designation")}${(v.idno || "").trim() ? " (ID: " + v.idno.trim() + ")" : ""}, to act for and on behalf of the Company for the following purposes:`));
      out.push(I(va(v, "scope", "scope of authority")));
      out.push(P(`This authority is given${v.res ? " under the resolution of the Board of Directors dated " + fdate(v.res) : ""} and is valid from ${vd(v, "from", "start date")} to ${vd(v, "upto", "end date")}, unless withdrawn earlier by written notice. The specimen signature of the authorised person is given below for your records.`));
      out.push(S(), L("Specimen signature of the authorised person: ____________________"), S());
      out.push(L("Yours faithfully,"), R("For **" + va(v, "company", "Company name") + "**"), S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")));
      return out;
    },
  },
  {
    id: "sharecert", group: "Share documents", title: "Share certificate (Form SH-1 layout)", short: "Share certificate in the layout of Form SH-1, ready to print and sign.",
    basis: "Section 46, Companies Act, 2013; Rule 5, Companies (Share Capital and Debentures) Rules, 2014 (Form SH-1)",
    fields: [
      { id: "certno", label: "Certificate number" }, { id: "folio", label: "Folio number" }, { id: "holder", label: "Name(s) of the holder(s)", type: "textarea" }, { id: "shares", label: "Number of shares", type: "number" },
      { id: "class", label: "Class of shares", def: "Equity" }, { id: "fv", label: "Face value per share (Rs.)", type: "number" }, { id: "paid", label: "Amount paid up per share (Rs.)", type: "number" },
      { id: "from", label: "Distinctive number from", type: "number" }, { id: "to", label: "Distinctive number to", type: "number" }, { id: "idate", label: "Date of issue", type: "date" },
      { id: "d1", label: "First director signing" }, { id: "d2", label: "Second director signing" }, { id: "cs", label: "Company Secretary or authorised officer" },
    ],
    notes: ["Issue the certificate within two months of incorporation or allotment, and within one month of receiving a transfer request (Section 56(4)).", "It must be signed by two directors and the Company Secretary, or by the authorised officer if there is no Company Secretary.", "Record the certificate in the Register of Members and the Register of Share Certificates (Form SH-2). Check the layout against the current Form SH-1 in the Rules before issue."],
    build(v) {
      const out = [C("**Form No. SH-1**"), C("[Pursuant to section 46(1) and rule 5(2) of the Companies (Share Capital and Debentures) Rules, 2014]"), S()].concat(header(v));
      out.push(T("SHARE CERTIFICATE"));
      out.push(L("Certificate No.: " + va(v, "certno", "number") + "          Folio No.: " + va(v, "folio", "folio")));
      out.push(P(`This is to certify that the person(s) named below is/are the registered holder(s) of the fully or partly paid shares of **${va(v, "company", "Company name")}**, as set out below, subject to the Memorandum and Articles of Association of the Company.`));
      out.push(L("Name(s) of the holder(s): **" + va(v, "holder", "name").replace(/\n+/g, ", ") + "**"));
      out.push(L("Number of shares: **" + vnum(v, "shares", "number of shares") + "**"));
      out.push(L("Class of shares: " + va(v, "class", "class")));
      out.push(L("Face value per share: Rs. " + va(v, "fv", "face value") + "/-"));
      out.push(L("Amount paid up per share: Rs. " + va(v, "paid", "paid-up amount") + "/-"));
      out.push(L("Distinctive numbers: from " + va(v, "from", "from") + " to " + va(v, "to", "to") + " (both inclusive)"));
      out.push(S(), P(`Given under the authority of the Board of Directors on ${vd(v, "idate", "date of issue")}.`));
      out.push(S(), S(), L("______________________________"), L("**" + va(v, "d1", "Director") + "**, Director"), S(), S(), L("______________________________"), L("**" + va(v, "d2", "Director") + "**, Director"), S(), S(), L("______________________________"), L("**" + va(v, "cs", "Company Secretary") + "**, Company Secretary / Authorised Officer"));
      return out;
    },
  },
];

/* ---------- more generators ---------- */
const LN = id => v => (v[id] || "").split(/\n+/).map(x => x.trim()).filter(Boolean);
const NUMLIST = (out, arr, ph) => (arr.length ? arr : [ph]).forEach((x, i) => out.push(I(`${i + 1}. ${x}`)));
const CO = v => va(v, "company", "Company name");
const MEETDATE = (v, id) => vd(v, id || "mdate", "date");
function bizItems(v, id, startAt, out) {
  let n = startAt;
  LN(id)(v).forEach(line => { const a = line.split("|"); const t = a[0].trim(); const k = /special/i.test(a[1] || "") ? "Special" : "Ordinary"; out.push(P(`**${n++}.** To consider and, if thought fit, to pass the following resolution as an ${k} Resolution: ${t}`)); });
  return n;
}
const GAP = "[Insert the reasons, the interest of the directors and key managerial personnel, and the other facts the members need to decide.]";

GENS.push(
  {
    id: "firstbm", group: "Meetings", title: "Minutes of the first Board meeting", short: "Minutes for a new private company's first Board meeting, with the usual first agenda.",
    basis: "Sections 173(2), 139(6), 184, 56(4) and 12, Companies Act, 2013; Secretarial Standard 1",
    fields: [
      { id: "idate", label: "Date of incorporation", type: "date" }, { id: "mdate", label: "Date of meeting", type: "date" }, { id: "mtime", label: "Time", ph: "11:00 a.m." }, { id: "place", label: "Place", ph: "Registered office" },
      { id: "chair", label: "Director elected to chair the meeting" }, { id: "present", label: "Directors present (one per line)", type: "textarea" },
      { id: "fy", label: "First financial year ends on", type: "date" }, { id: "firm", label: "First auditor: firm name" }, { id: "frn", label: "Firm registration number" }, { id: "remun", label: "Auditor's remuneration (Rs.)", type: "number" },
      { id: "bank", label: "Bank for the company's account" }, { id: "branch", label: "Branch" }, { id: "signs", label: "Authorised bank signatories (one per line)", type: "textarea" },
      { id: "holders", label: "Subscribers and their shares, for certificates (one per line)", type: "textarea", ph: "Mr. A: 5,000 equity shares" }, { id: "exp", label: "Preliminary expenses to be ratified (Rs.)", type: "number" }, { id: "sdate", label: "Date of signing the minutes", type: "date" },
    ],
    notes: ["The first Board meeting must be held within 30 days of incorporation (Section 173(2)).", "The first auditor must be appointed by the Board within 30 days of incorporation, failing which the members must appoint at an EGM within 90 days (Section 139(6)).", "Each director gives a notice of interest in Form MBP-1 (Section 184) at the first meeting. Declaration of commencement of business is filed in Form INC-20A within 180 days of incorporation (Section 10A)."],
    build(v) {
      const out = header(v);
      out.push(T(`MINUTES OF THE FIRST MEETING OF THE BOARD OF DIRECTORS HELD ON ${MEETDATE(v).toUpperCase()}${v.mtime ? " AT " + v.mtime.toUpperCase() : ""} AT ${PLACE(v).toUpperCase()}`));
      out.push(P("**Directors present**")); NUMLIST(out, LN("present")(v), "[names of directors]");
      out.push(S(), P(`The directors present elected ${va(v, "chair", "name")} as Chairman of the meeting. The Chairman noted that the meeting is held within thirty days of the incorporation of the Company on ${vd(v, "idate", "date of incorporation")}, as required by Section 173(2) of the Companies Act, 2013, that all the directors had been given notice, and that the required quorum was present.`));
      let n = 1; const item = (h, ...ps) => { out.push(P(`**${n++}. ${h}**`)); ps.forEach(x => out.push(x)); };
      item("Certificate of incorporation", P("The Chairman placed before the Board the Certificate of Incorporation issued by the Registrar of Companies, together with the Memorandum and Articles of Association of the Company, and the Board took note of them."));
      item("Registered office", RES(`the registered office of the Company at ${va(v, "ro", "address")} be and is hereby confirmed, and a notice of the situation of the registered office be filed in Form INC-22 within thirty days of incorporation.`));
      item("Disclosure of interest", P("Every director present gave a notice of disclosure of interest in Form MBP-1 under Section 184(1), together with the declaration and consent in Forms DIR-8 and DIR-2. The notices were taken on record and will be kept in the records of the Company."));
      item("First auditor", RES(`pursuant to Section 139(6) of the Companies Act, 2013, ${va(v, "firm", "name of firm")}, Chartered Accountants (Firm Registration No. ${va(v, "frn", "FRN")}), be and are hereby appointed as the first auditor of the Company to hold office until the conclusion of the first Annual General Meeting, at a remuneration of ${vmoney(v, "remun", "remuneration")} plus applicable taxes, and Form ADT-1 be filed within the prescribed time.`));
      item("Financial year", RES(`the financial year of the Company shall end on ${vd(v, "fy", "date")} of each year, and the first financial year shall end on ${vd(v, "fy", "date")}.`));
      item("Bank account", RES(`a current account of the Company be opened with ${va(v, "bank", "bank")}, ${va(v, "branch", "branch")} Branch, and the following persons be authorised to operate it:`));
      NUMLIST(out, LN("signs")(v), "[name of signatory]");
      item("Share certificates", RES(`share certificates in Form SH-1 be issued to the subscribers to the Memorandum within two months of incorporation, as under:`));
      NUMLIST(out, LN("holders")(v), "[subscriber and shares]");
      item("Preliminary expenses", RES(`the expenses incurred in connection with the incorporation of the Company, amounting to ${vmoney(v, "exp", "amount")}, be ratified and paid by the Company.`));
      item("Statutory registers and records", RES("the Registers required by the Companies Act, 2013, including the Register of Members, the Register of Directors and Key Managerial Personnel and the Register of Contracts, be opened and kept at the registered office, and the Secretarial Standards be followed."));
      item("Commencement of business", P("The Board noted that the subscribers have paid for their shares and that a declaration of commencement of business in Form INC-20A has to be filed within one hundred and eighty days of incorporation."));
      out.push(S(), P("There being no other business, the meeting ended with a vote of thanks to the Chair."));
      return out.concat([S(), S(), R("**" + va(v, "chair", "Chairman") + "**"), R("Chairman"), L("Place: " + (v.place || "")), L("Date of signing: " + fdate(v.sdate))]);
    },
  },
  {
    id: "agmmin", group: "Meetings", title: "Minutes of Annual General Meeting", short: "Minutes with attendance, quorum, each item and the result of the vote.",
    basis: "Section 118, Companies Act, 2013; Secretarial Standard 2",
    fields: [
      { id: "no", label: "Number of the AGM", ph: "5th" }, { id: "mdate", label: "Date of meeting", type: "date" }, { id: "start", label: "Started at", ph: "11:00 a.m." }, { id: "end", label: "Ended at", ph: "11:45 a.m." }, { id: "place", label: "Place" },
      { id: "chair", label: "Chairman of the meeting" }, { id: "members", label: "Number of members present in person or by proxy", type: "number" }, { id: "dirs", label: "Directors present (one per line)", type: "textarea" },
      { id: "auditor", label: "Auditor or representative present" }, { id: "fyend", label: "Financial year ended on", type: "date" }, { id: "retire", label: "Director retiring by rotation and reappointed" },
      { id: "special", label: "Special business (one per line, then | Ordinary or | Special)", type: "textarea" }, { id: "vote", label: "How the votes were taken", type: "select", opts: [["by show of hands", "By show of hands"], ["by e-voting and poll", "By e-voting and poll"]] },
      { id: "sdate", label: "Date of signing", type: "date" },
    ],
    notes: ["Quorum for a private company is two members present in person. For a public company it depends on the number of members (Section 103).", "Record the results of the vote on each item. Where voting was by poll, attach the scrutiniser's report.", "Minutes of general meetings should be signed within 30 days and each page initialled (Secretarial Standard 2)."],
    build(v) {
      const out = header(v); let n = 1;
      out.push(T(`MINUTES OF THE ${va(v, "no", "number").toUpperCase()} ANNUAL GENERAL MEETING HELD ON ${MEETDATE(v).toUpperCase()} AT ${PLACE(v).toUpperCase()}`));
      out.push(P("**Directors present**")); NUMLIST(out, LN("dirs")(v), "[names]");
      out.push(L(`Members present in person or by proxy: ${va(v, "members", "number")}`)); if (v.auditor) out.push(L("Auditor present: " + v.auditor));
      out.push(S(), P(`${va(v, "chair", "name")} took the chair and, the required quorum being present, called the meeting to order at ${va(v, "start", "time")}. With the consent of the members the notice convening the meeting was taken as read. The Chairman declared that the voting on all items would be taken ${v.vote || "by show of hands"}.`));
      out.push(P(`**${n++}. Adoption of financial statements.** The audited financial statements for the year ended ${vd(v, "fyend", "year end")}, with the reports of the Board and the Auditors, were taken as read. It was proposed and seconded and was passed as an Ordinary Resolution that they be adopted.`));
      if ((v.retire || "").trim()) out.push(P(`**${n++}. Reappointment of a director.** ${v.retire.trim()} retired by rotation under Section 152(6) and, being eligible, was reappointed as a director liable to retire by rotation, by an Ordinary Resolution.`));
      LN("special")(v).forEach(line => { const a = line.split("|"); const k = /special/i.test(a[1] || "") ? "Special" : "Ordinary"; out.push(P(`**${n++}.** ${a[0].trim()}. The resolution was proposed and seconded and was passed as a ${k} Resolution by the requisite majority.`)); });
      out.push(S(), P(`There being no other business, the meeting ended at ${va(v, "end", "time")} with a vote of thanks to the Chair.`));
      return out.concat([S(), S(), R("**" + va(v, "chair", "Chairman") + "**"), R("Chairman"), L("Place: " + (v.place || "")), L("Date of signing: " + fdate(v.sdate))]);
    },
  },
  {
    id: "egmnotice", group: "Meetings", title: "Notice of Extraordinary General Meeting", short: "EGM notice with the explanatory statement for each item.",
    basis: "Sections 100, 101, 102 and 105, Companies Act, 2013; Secretarial Standard 2",
    fields: [
      { id: "ndate", label: "Date of notice", type: "date" }, { id: "mdate", label: "Date of EGM", type: "date" }, { id: "mtime", label: "Time", ph: "11:00 a.m." }, { id: "venue", label: "Venue", type: "textarea" },
      { id: "biz", label: "Business (one per line, then | Ordinary or | Special)", type: "textarea", ph: "To increase the authorised share capital | Ordinary" },
      { id: "short", label: "Held at shorter notice?", type: "select", opts: [["no", "No (21 clear days)"], ["yes", "Yes (consent of 95% of members)"]] },
      { id: "sname", label: "Name of the person issuing the notice" }, { id: "sdes", label: "Designation" },
    ],
    notes: ["Send at least 21 clear days before the meeting, unless members holding at least 95% of the voting power consent in writing to a shorter notice.", "Attach an explanatory statement for every item (Section 102). Fill in the placeholders.", "Attach the proxy form (Form MGT-11) and an attendance slip."],
    build(v) {
      const out = header(v); const biz = LN("biz")(v);
      out.push(T("NOTICE OF EXTRAORDINARY GENERAL MEETING")); out.push(R("Date: " + vd(v, "ndate", "date of notice")));
      out.push(P(`**NOTICE** is hereby given that an Extraordinary General Meeting of the members of **${CO(v)}** will be held on **${vd(v, "mdate", "date")}** at **${va(v, "mtime", "time")}** at ${va(v, "venue", "venue")}${v.short === "yes" ? ", at shorter notice with the consent of the members" : ""}, to transact the following business:`));
      out.push(P("**SPECIAL BUSINESS**")); bizItems(v, "biz", 1, out);
      out.push(S(), L("By order of the Board"), R("For **" + CO(v) + "**"), S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")));
      out.push(S(), P("**Notes**"), P("1. A member entitled to attend and vote may appoint a proxy, who need not be a member. The proxy form must reach the registered office at least 48 hours before the meeting."), P("2. An explanatory statement under Section 102 is annexed."));
      out.push(S(), P("**EXPLANATORY STATEMENT UNDER SECTION 102 OF THE COMPANIES ACT, 2013**"));
      biz.forEach((b, i) => { out.push(P(`**Item No. ${i + 1}:** ${b.split("|")[0].trim()}`)); out.push(P(GAP)); });
      return out;
    },
  },
  {
    id: "proxy", group: "Meetings", title: "Proxy form and attendance slip (MGT-11)", short: "Proxy form in the layout of Form MGT-11, with an attendance slip.",
    basis: "Section 105(6), Companies Act, 2013; Rule 19(3), Companies (Management and Administration) Rules, 2014",
    fields: [
      { id: "kind", label: "Meeting", type: "select", opts: [["Annual General Meeting", "Annual General Meeting"], ["Extraordinary General Meeting", "Extraordinary General Meeting"]] }, { id: "mdate", label: "Date of meeting", type: "date" },
      { id: "mtime", label: "Time", ph: "11:00 a.m." }, { id: "venue", label: "Venue", type: "textarea" }, { id: "res", label: "Resolutions (one per line)", type: "textarea", def: "Adoption of financial statements\nReappointment of a director retiring by rotation" },
    ],
    notes: ["The proxy must be deposited at the registered office not less than 48 hours before the meeting.", "A proxy can vote only on a poll. Check the current Form MGT-11 in the Rules before sending it out."],
    build(v) {
      const out = [C("**Form No. MGT-11**"), C("**PROXY FORM**"), C("[Pursuant to section 105(6) of the Companies Act, 2013 and rule 19(3) of the Companies (Management and Administration) Rules, 2014]"), S()].concat(header(v));
      out.push(L("Name of the member(s): ______________________________"), L("Registered address: ______________________________"), L("E-mail ID: ______________________   Folio No. / Client ID: ____________   DP ID: ____________"), S());
      out.push(P(`I/We, being the member(s) of ____________ shares of the above-named company, hereby appoint:`));
      out.push(I("1. Name: ____________________  Address: ____________________  E-mail ID: ______________  Signature: ____________, or failing him/her"), I("2. Name: ____________________  Address: ____________________  E-mail ID: ______________  Signature: ____________, or failing him/her"), I("3. Name: ____________________  Address: ____________________  E-mail ID: ______________  Signature: ____________"));
      out.push(P(`as my/our proxy to attend and vote (on a poll) for me/us and on my/our behalf at the ${va(v, "kind", "meeting")} of the Company, to be held on ${vd(v, "mdate", "date")} at ${va(v, "mtime", "time")} at ${va(v, "venue", "venue")} and at any adjournment of it in respect of the resolutions indicated below:`));
      out.push(L("**Resolutions** (tick For or Against against each)"));
      LN("res")(v).forEach((r, i) => { out.push(L(`${i + 1}. ${r}`)); out.push(I("For [   ]        Against [   ]")); });
      out.push(S(), L("Signed this ______ day of ____________ 20____"), S(), L("Signature of shareholder: ____________________"), L("Signature of proxy holder(s): ____________________"));
      out.push(S(), P("**Note:** This form of proxy, to be effective, should be duly completed and deposited at the registered office of the Company not less than 48 hours before the commencement of the meeting."));
      out.push(PB(), C("**ATTENDANCE SLIP**"), S(), ...header(v), L("Folio No. / DP ID and Client ID: ____________________"), L("Name and address of the member: ____________________"), L("Number of shares held: ____________"), S(),
        P(`I hereby record my presence at the ${va(v, "kind", "meeting")} of the Company held on ${vd(v, "mdate", "date")} at ${va(v, "mtime", "time")} at ${va(v, "venue", "venue")}.`), S(), L("Signature of the member / proxy: ____________________"));
      return out;
    },
  },
  {
    id: "dispatch", group: "Meetings", title: "Proof of dispatch of notice", short: "A certificate that the notice of a meeting was sent, with the list of people it went to.",
    basis: "Sections 101 and 173(3), Companies Act, 2013; Secretarial Standards 1 and 2",
    fields: [
      { id: "kind", label: "Meeting", type: "select", opts: [["Board meeting", "Board meeting"], ["Committee meeting", "Committee meeting"], ["Annual General Meeting", "Annual General Meeting"], ["Extraordinary General Meeting", "Extraordinary General Meeting"]] },
      { id: "mdate", label: "Date of meeting", type: "date" }, { id: "ddate", label: "Date of dispatch", type: "date" }, { id: "mode", label: "Mode of sending", ph: "e-mail / registered post / hand delivery" },
      { id: "to", label: "Sent to (one per line)", type: "textarea" }, { id: "sname", label: "Name of the person certifying" }, { id: "sdes", label: "Designation", ph: "Company Secretary" },
    ],
    notes: ["Keep this with the notice in the minute book. For a Board meeting, notice must be sent at least seven days before the meeting."],
    build(v) {
      const out = header(v);
      out.push(T("CERTIFICATE OF DISPATCH OF NOTICE"));
      out.push(P(`I certify that the notice of the ${va(v, "kind", "meeting")} of **${CO(v)}** to be held on ${vd(v, "mdate", "meeting date")}, with the agenda and the notes to the agenda or the explanatory statement, was sent on ${vd(v, "ddate", "dispatch date")} by ${va(v, "mode", "mode")} to the following persons, at the address or the e-mail address they have given to the Company:`));
      NUMLIST(out, LN("to")(v), "[name of recipient]");
      return out.concat([S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")), L("Date: " + vd(v, "ddate", "date"))]);
    },
  },
  {
    id: "declare", group: "Directors", title: "Director's declarations (MBP-1, DIR-8, DIR-2)", short: "Disclosure of interest, declaration of non-disqualification and consent to act.",
    basis: "Sections 184, 164, 165 and 152, Companies Act, 2013; Rules 9, 14 and 8 of the Companies Rules",
    fields: [
      { id: "kind", label: "Document", type: "select", opts: [["mbp1", "Notice of interest (MBP-1)"], ["dir8", "Declaration of non-disqualification (DIR-8)"], ["dir2", "Consent to act as a director (DIR-2)"]] },
      { id: "name", label: "Name of the director" }, { id: "din", label: "DIN" }, { id: "addr", label: "Residential address", type: "textarea" }, { id: "date", label: "Date", type: "date" }, { id: "place", label: "Place" },
      { id: "desig", label: "Position being taken", ph: "Additional Director", show: v => v.kind !== "mbp1" }, { id: "effdate", label: "Date of appointment", type: "date", show: v => v.kind !== "mbp1" },
      { id: "ints", label: "Companies, firms or bodies where you are interested (one per line: name | nature of interest | shareholding)", type: "textarea", show: v => (v.kind || "mbp1") === "mbp1" },
    ],
    notes: ["MBP-1 is given at the first Board meeting of each financial year and again whenever the interest changes (Section 184(1)).", "DIR-8 and DIR-2 are attached to the appointment filing. Use the current form on the MCA portal where the Rules require an e-form."],
    build(v) {
      const out = header(v), k = v.kind || "mbp1", nm = va(v, "name", "name"), din = va(v, "din", "DIN");
      if (k === "mbp1") {
        out.push(C("**Form No. MBP-1**"), C("**NOTICE OF INTEREST BY DIRECTOR**"), C("[Pursuant to section 184(1) and rule 9(1) of the Companies (Meetings of Board and its Powers) Rules, 2014]"), S());
        out.push(P(`I, ${nm}, son/daughter/spouse of ____________, a director of **${CO(v)}**, resident of ${va(v, "addr", "address").replace(/\n+/g, ", ")} (DIN: ${din}), give notice that I am interested in the following companies, bodies corporate, firms or other associations of individuals:`));
        const rows = LN("ints")(v); (rows.length ? rows : ["[name] | [nature of interest] | [shareholding]"]).forEach((r, i) => { const a = r.split("|").map(x => x.trim()); out.push(I(`${i + 1}. Name: ${a[0] || ""}. Nature of interest or concern: ${a[1] || ""}. Shareholding: ${a[2] || ""}.`)); });
        out.push(P("I shall give further notice to the Company of any change in these particulars, within the time the Act allows, and in any case at the first Board meeting of every financial year."));
      } else if (k === "dir8") {
        out.push(C("**Form DIR-8**"), C("**DECLARATION BY A DIRECTOR**"), C("[Pursuant to rule 14 of the Companies (Appointment and Qualification of Directors) Rules, 2014]"), S());
        out.push(P(`I, ${nm} (DIN: ${din}), resident of ${va(v, "addr", "address").replace(/\n+/g, ", ")}, proposed to be appointed as ${va(v, "desig", "position")} of **${CO(v)}** with effect from ${vd(v, "effdate", "date")}, declare that:`));
        out.push(I("1. I am not disqualified from being appointed as a director under any of the grounds in Section 164(1) of the Companies Act, 2013."), I("2. I have not been a director of any company that has defaulted in the ways listed in Section 164(2), so as to attract a disqualification for me."), I("3. The number of directorships I hold, including this one, is within the limits of Section 165."), I("4. I have not been removed from the office of a director by any court or tribunal, and I have not been convicted as described in Section 164."));
        out.push(P("I undertake to inform the Company at once if any of the above changes."));
      } else {
        out.push(C("**CONSENT TO ACT AS A DIRECTOR**"), C("[Pursuant to section 152(5) of the Companies Act, 2013 and rule 8 of the Companies (Appointment and Qualification of Directors) Rules, 2014]"), S());
        out.push(P(`I, ${nm} (DIN: ${din}), resident of ${va(v, "addr", "address").replace(/\n+/g, ", ")}, hereby give my consent to act as ${va(v, "desig", "position")} of **${CO(v)}** with effect from ${vd(v, "effdate", "date")}. I confirm that I hold a valid Director Identification Number, that I am not disqualified under Section 164, and that I will comply with the duties of a director under Section 166.`));
      }
      return out.concat([S(), S(), L("Place: " + (v.place || "")), L("Date: " + vd(v, "date", "date")), R("**" + nm + "**"), R("(DIN: " + din + ")")]);
    },
  },
  {
    id: "boardreport", group: "Annual filing", title: "Board's report (outline)", short: "The headings a Board's report must cover, filled with your figures.",
    basis: "Section 134(3), (5) and Rule 8, Companies (Accounts) Rules, 2014",
    fields: [
      { id: "fy", label: "Financial year", ph: "2025-26" }, { id: "fyend", label: "Year ended on", type: "date" }, { id: "type", label: "Type of company", type: "select", opts: [["full", "Full report"], ["small", "Small company or One Person Company (abridged)"]] },
      { id: "rev", label: "Revenue (Rs.)", type: "number" }, { id: "profit", label: "Profit after tax (Rs.)", type: "number" }, { id: "div", label: "Dividend recommended", ph: "Nil" }, { id: "reserve", label: "Amount transferred to reserves", ph: "Nil" },
      { id: "nature", label: "Change in the nature of business", ph: "None" }, { id: "meets", label: "Number of Board meetings held", type: "number" }, { id: "dirchg", label: "Changes in directors and key managerial personnel", type: "textarea", ph: "None" },
      { id: "auditor", label: "Statutory auditor and term" }, { id: "loans", label: "Loans, guarantees and investments (Section 186)", ph: "None" }, { id: "rpt", label: "Related party transactions", ph: "None, so AOC-2 is not required" },
      { id: "matl", label: "Material changes after the year end", ph: "None" }, { id: "sub", label: "Subsidiaries, joint ventures and associates", ph: "None" }, { id: "dep", label: "Deposits", ph: "The Company has not accepted any deposits" },
      { id: "date", label: "Date of the report", type: "date" }, { id: "place", label: "Place" }, { id: "d1", label: "First director signing" }, { id: "d1din", label: "DIN" }, { id: "d2", label: "Second director signing" }, { id: "d2din", label: "DIN" },
    ],
    notes: ["Section 134(3) and Rule 8 list what the report must say. A small company or One Person Company may give a shorter report.", "Listed companies and large companies must add more (for example, the annual return web link, CSR, risk management and the remuneration details). Check each against the Act.", "Attach AOC-2 when there are related party contracts, and a statement on the employee and energy details where Rule 8(3) applies."],
    build(v) {
      const out = header(v), full = v.type !== "small"; let n = 1;
      out.push(T(`BOARD'S REPORT FOR THE FINANCIAL YEAR ${va(v, "fy", "year").toUpperCase()}`));
      out.push(P(`To the Members of **${CO(v)}**`)); out.push(P(`Your directors present their report on the business and operations of the Company for the year ended ${vd(v, "fyend", "year end")}.`));
      const h = (t, ...ps) => { out.push(P(`**${n++}. ${t}**`)); ps.forEach(x => out.push(P(x))); };
      h("Financial results", `Revenue: ${vmoney(v, "rev", "revenue")}. Profit after tax: ${vmoney(v, "profit", "profit")}.`);
      h("State of the Company's affairs", "[Describe the business, the performance in the year and the outlook.]");
      h("Dividend", `${va(v, "div", "dividend")}.`); h("Transfer to reserves", `${va(v, "reserve", "amount")}.`);
      h("Change in the nature of business", `${va(v, "nature", "change")}.`);
      h("Material changes after the year end", `${va(v, "matl", "changes")}. There are no orders of a regulator, court or tribunal that affect the Company's going concern status and future operations.`);
      h("Directors and key managerial personnel", va(v, "dirchg", "changes") + ".");
      h("Meetings of the Board", `The Board met ${va(v, "meets", "number")} times in the year, and the gap between any two meetings was within the limit in the Act.`);
      h("Directors' responsibility statement (Section 134(5))", "The directors confirm that: (a) the applicable accounting standards were followed and there are no material departures; (b) the accounting policies were applied consistently and the judgments made give a true and fair view; (c) proper and sufficient care was taken to keep adequate accounting records, to safeguard the assets and to prevent and detect fraud; (d) the annual accounts were prepared on a going concern basis; and (e) proper systems were devised to ensure compliance with all applicable laws, and they were adequate and operating effectively.");
      h("Auditors", `${va(v, "auditor", "auditor and term")}. The Auditors' Report has no qualification, reservation or adverse remark.`);
      h("Loans, guarantees and investments", `${va(v, "loans", "details")}.`);
      h("Related party transactions", `${va(v, "rpt", "details")}.`);
      h("Subsidiaries, joint ventures and associates", `${va(v, "sub", "details")}.`); h("Deposits", `${va(v, "dep", "details")}.`);
      h("Internal financial controls", "The Company has adequate internal financial controls for the size and nature of its business, and they were operating effectively in the year.");
      if (full) { h("Particulars required by Rule 8(3)", "Conservation of energy: [details]. Technology absorption: [details]. Foreign exchange earned and spent: [details]."); h("Risk management and corporate social responsibility", "[Describe the risk management policy. State whether Section 135 applies, and if it does, give the CSR details and the Annual Report on CSR.]"); h("Annual return", "The annual return is available at [web link of the Company], if the Company has a website."); }
      h("Compliance with Secretarial Standards, and other statements", "The Company has complied with the applicable Secretarial Standards. No frauds were reported by the auditors under Section 143(12). The Company has complied with the Sexual Harassment of Women at Workplace (Prevention, Prohibition and Redressal) Act, 2013, where applicable, and has constituted an Internal Committee where needed.");
      out.push(P("**Acknowledgement.** Your directors thank the bankers, customers, employees and members for their support."));
      return out.concat([S(), L("For and on behalf of the Board of Directors"), R("For **" + CO(v) + "**"), S(), S(), R("**" + va(v, "d1", "Director") + "** (DIN: " + va(v, "d1din", "DIN") + ")"), R("Director"), S(), R("**" + va(v, "d2", "Director") + "** (DIN: " + va(v, "d2din", "DIN") + ")"), R("Director"), L("Place: " + (v.place || "")), L("Date: " + fdate(v.date))]);
    },
  },
  {
    id: "aoc2", group: "Annual filing", title: "Form AOC-2 (related party contracts)", short: "Particulars of contracts with related parties, in the layout of Form AOC-2.",
    basis: "Section 188(1) and Rule 8(2), Companies (Accounts) Rules, 2014",
    fields: [
      { id: "fy", label: "Financial year", ph: "2025-26" },
      { id: "n1", label: "Not at arm's length: name of related party and the relationship", type: "textarea" }, { id: "n2", label: "Nature of the contract", type: "textarea" }, { id: "n3", label: "Duration" }, { id: "n4", label: "Main terms, including the value", type: "textarea" },
      { id: "n5", label: "Justification for entering into it", type: "textarea" }, { id: "n6", label: "Date of Board approval", type: "date" }, { id: "n7", label: "Amount paid as advance, if any" }, { id: "n8", label: "Date of the special resolution under the first proviso to Section 188(1)", type: "date" },
      { id: "a1", label: "At arm's length: name of related party and the relationship", type: "textarea" }, { id: "a2", label: "Nature of the contract", type: "textarea" }, { id: "a3", label: "Duration" }, { id: "a4", label: "Main terms, including the value", type: "textarea" }, { id: "a5", label: "Date of Board approval", type: "date" }, { id: "a6", label: "Amount paid as advance, if any" },
      { id: "sname", label: "Name of the signing director" }, { id: "sdin", label: "DIN" }, { id: "place", label: "Place" }, { id: "date", label: "Date", type: "date" },
    ],
    notes: ["Write NIL where there are no such contracts. Attach the form to the Board's report.", "Check the current Form AOC-2 in the Rules for the exact wording of each heading."],
    build(v) {
      const nil = id => { const x = (v[id] || "").trim(); return x ? x.replace(/\n+/g, "; ") : "NIL"; };
      const out = [C("**Form No. AOC-2**"), C("[Pursuant to clause (h) of sub-section (3) of section 134 of the Act and Rule 8(2) of the Companies (Accounts) Rules, 2014]"), C("**Form for disclosure of particulars of contracts/arrangements entered into by the company with related parties referred to in sub-section (1) of section 188 of the Companies Act, 2013 including certain arm's length transactions under third proviso thereto**"), S(), L(`Company: ${CO(v)}.   Financial year: ${va(v, "fy", "year")}`), S()];
      out.push(P("**1. Details of contracts or arrangements or transactions not at arm's length basis**"));
      [["(a) Name(s) of the related party and nature of relationship", "n1"], ["(b) Nature of contracts/arrangements/transactions", "n2"], ["(c) Duration of the contracts/arrangements/transactions", "n3"], ["(d) Salient terms, including the value, if any", "n4"], ["(e) Justification for entering into such contracts or arrangements or transactions", "n5"], ["(f) Date(s) of approval by the Board", "n6"], ["(g) Amount paid as advances, if any", "n7"], ["(h) Date on which the special resolution was passed in general meeting as required under first proviso to section 188", "n8"]].forEach(a => out.push(I(`${a[0]}: ${a[1] === "n6" || a[1] === "n8" ? ((v[a[1]] || "").trim() ? fdate(v[a[1]]) : "NIL") : nil(a[1])}`)));
      out.push(S(), P("**2. Details of material contracts or arrangement or transactions at arm's length basis**"));
      [["(a) Name(s) of the related party and nature of relationship", "a1"], ["(b) Nature of contracts/arrangements/transactions", "a2"], ["(c) Duration of the contracts/arrangements/transactions", "a3"], ["(d) Salient terms, including the value, if any", "a4"], ["(e) Date(s) of approval by the Board, if any", "a5"], ["(f) Amount paid as advances, if any", "a6"]].forEach(a => out.push(I(`${a[0]}: ${a[1] === "a5" ? ((v[a[1]] || "").trim() ? fdate(v[a[1]]) : "NIL") : nil(a[1])}`)));
      return out.concat([S(), S(), R("For and on behalf of the Board of Directors"), R("**" + va(v, "sname", "name") + "**"), R("Director (DIN: " + va(v, "sdin", "DIN") + ")"), L("Place: " + (v.place || "")), L("Date: " + fdate(v.date))]);
    },
  },
  {
    id: "dupcert", group: "Share documents", title: "Duplicate share certificate resolution", short: "Board resolution to issue a duplicate or renewed share certificate.",
    basis: "Section 46(2), Companies Act, 2013; Rule 6, Companies (Share Capital and Debentures) Rules, 2014",
    fields: [
      { id: "mdate", label: "Date of Board meeting", type: "date" }, { id: "holder", label: "Name of the holder" }, { id: "folio", label: "Folio number" }, { id: "oldno", label: "Original certificate number" }, { id: "shares", label: "Number of shares", type: "number" },
      { id: "dist", label: "Distinctive numbers", ph: "101 to 600" }, { id: "reason", label: "Reason", type: "select", opts: [["lost", "Lost"], ["destroyed", "Destroyed"], ["defaced", "Torn, defaced or mutilated"], ["full", "Pages for transfer are used up"]] },
      { id: "docs", label: "Documents received", ph: "application, FIR copy, affidavit and indemnity bond" }, { id: "newno", label: "New certificate number" }, { id: "auth", label: "Person authorised (name and designation)" }, { id: "sname", label: "Name of the signing director" },
    ],
    notes: ["Check the Articles and Rule 6 for the documents the company must take and any public notice for a lost certificate.", "The new certificate must be marked 'Duplicate', and the Register of Renewed and Duplicate Share Certificates (Form SH-2) must be updated."],
    build(v) {
      const out = header(v); out.push(T("RESOLUTION PASSED AT THE MEETING OF THE BOARD OF DIRECTORS HELD ON " + vd(v, "mdate", "date").toUpperCase()));
      const why = { lost: "has been lost", destroyed: "has been destroyed", defaced: "has become torn, defaced or mutilated", full: "has no space left for transfer endorsements" }[v.reason || "lost"];
      out.push(P("**Issue of a duplicate share certificate**"));
      out.push(RES(`the Board considered the request of ${va(v, "holder", "name")} (Folio No. ${va(v, "folio", "folio")}) that Share Certificate No. ${va(v, "oldno", "number")} for ${vnum(v, "shares", "number")} shares (distinctive numbers ${va(v, "dist", "numbers")}) ${why}, and the documents received, being ${va(v, "docs", "documents")}, and, pursuant to Section 46(2) of the Companies Act, 2013 and Rule 6 of the Companies (Share Capital and Debentures) Rules, 2014, approved the issue of a duplicate certificate in its place.`));
      out.push(RESF(`a duplicate certificate numbered ${va(v, "newno", "number")}, marked "Duplicate", be issued on the same terms, the original certificate (if any) be cancelled, and the details be entered in the Register of Renewed and Duplicate Share Certificates (Form SH-2) and the Register of Members.`));
      out.push(RESF(`${va(v, "auth", "authorised person")} be authorised to sign and issue the certificate and to do all things needed to give effect to this resolution.`));
      return out.concat(signoff(v, va(v, "sname", "name"), "Director"));
    },
  },
  {
    id: "auditorkit", group: "Annual filing", title: "Auditor appointment kit", short: "Board resolution, request letter, consent and eligibility certificate, and the AGM resolution.",
    basis: "Sections 139(1), 141 and 142, Companies Act, 2013; Rule 4, Companies (Audit and Auditors) Rules, 2014",
    fields: [
      { id: "mdate", label: "Date of Board meeting", type: "date" }, { id: "firm", label: "Auditor's firm name" }, { id: "frn", label: "Firm registration number" }, { id: "faddr", label: "Firm's address", type: "textarea" },
      { id: "agmno", label: "Number of the AGM", ph: "5th" }, { id: "agmdate", label: "Date of AGM", type: "date" }, { id: "term", label: "Concluding at the AGM for the year", ph: "2030-31" }, { id: "remun", label: "Remuneration (Rs.)", type: "number" },
      { id: "sname", label: "Name of the signing director" }, { id: "sdes", label: "Designation", ph: "Director" },
    ],
    notes: ["An auditor is appointed for five years, until the sixth AGM, by the members. Form ADT-1 is filed within 15 days of the AGM.", "The auditor must give written consent and a certificate that the appointment is within the limits of Section 141, and give the details the Rules require.", "A listed company, or one in the classes in the Rules, must rotate its auditors (Section 139(2)). Check that."],
    build(v) {
      const out = []; const f = va(v, "firm", "firm");
      out.push(...header(v), T("1. RESOLUTION PASSED AT THE MEETING OF THE BOARD OF DIRECTORS HELD ON " + vd(v, "mdate", "date").toUpperCase()));
      out.push(RES(`the Board recommends to the members that ${f}, Chartered Accountants (Firm Registration No. ${va(v, "frn", "FRN")}), be appointed as the statutory auditor of the Company under Section 139(1) of the Companies Act, 2013, to hold office from the conclusion of the ${va(v, "agmno", "number")} Annual General Meeting until the conclusion of the sixth Annual General Meeting after it, being the meeting for the financial year ${va(v, "term", "year")}, at a remuneration of ${vmoney(v, "remun", "amount")} plus applicable taxes and out-of-pocket expenses, and subject to the written consent and eligibility certificate of the auditor.`));
      out.push(...signoff(v, va(v, "sname", "name"), va(v, "sdes", "designation")), PB());
      out.push(...header(v), R("Date: [date]"), L("To,"), L(f), L(va(v, "faddr", "address").replace(/\n+/g, ", ")), S(), P("**Subject: Proposal to appoint you as the statutory auditor**"));
      out.push(P(`The Board has proposed to the members that you be appointed as the statutory auditor of the Company at the ${va(v, "agmno", "number")} Annual General Meeting to be held on ${vd(v, "agmdate", "date")}. Under Section 139(1) and Rule 4 you are requested to give your written consent and a certificate that you satisfy the criteria in Section 141 and that your appointment is within the limits of Section 141(3)(g). Please confirm the number of audits you hold.`));
      out.push(L("Yours faithfully,"), R("For **" + CO(v) + "**"), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")), PB());
      out.push(C("**CONSENT AND CERTIFICATE OF ELIGIBILITY**"), C("[To be issued by the auditor on its own letterhead]"), S(), L("To the Board of Directors, " + CO(v)), S());
      out.push(P(`We, ${f}, Chartered Accountants (FRN: ${va(v, "frn", "FRN")}), give our consent to be appointed as the statutory auditor of **${CO(v)}** under Section 139(1) of the Companies Act, 2013. We certify that: (a) we are eligible for appointment and are not disqualified under Section 141; (b) the appointment, if made, will be within the limits laid down in Section 141(3)(g); (c) our firm holds a valid certificate from the Peer Review Board of the Institute of Chartered Accountants of India, where it applies; and (d) there is no pending proceeding or order against us that affects this appointment.`));
      out.push(S(), L("For " + f), L("Partner / Proprietor: ____________________"), L("Membership No.: ____________"), L("Date: ____________"), PB());
      out.push(...header(v), T("ITEM FOR THE NOTICE OF THE " + va(v, "agmno", "number").toUpperCase() + " ANNUAL GENERAL MEETING"));
      out.push(P(`To consider and, if thought fit, to pass the following resolution as an Ordinary Resolution: **RESOLVED THAT** pursuant to Sections 139, 142 and other applicable provisions of the Companies Act, 2013 and the Companies (Audit and Auditors) Rules, 2014, ${f} (Firm Registration No. ${va(v, "frn", "FRN")}) be and is hereby appointed as the statutory auditor of the Company to hold office from the conclusion of this meeting until the conclusion of the sixth Annual General Meeting after it, at a remuneration fixed by the Board, in consultation with the auditor.`));
      out.push(P("After the AGM, file Form ADT-1 with the Registrar of Companies within fifteen days."));
      return out;
    },
  },
);

/* policies ------------------------------------------------------------------ */
const POL_FIELDS = [{ id: "adopt", label: "Date of Board approval", type: "date" }, { id: "eff", label: "Effective from", type: "date" }, { id: "contact", label: "Name of the contact person", ph: "Company Secretary" }, { id: "email", label: "E-mail of the contact person" }];
function polHead(v, title, basis) { const o = header(v); o.push(T(title)); o.push(C(basis)); o.push(L(`Approved by the Board on ${vd(v, "adopt", "date")}. Effective from ${vd(v, "eff", "date")}.`), S()); return o; }
function clause(o, n, h, t) { o.push(P(`**${n}. ${h}.** ${t}`)); }
GENS.push(
  {
    id: "vigil", group: "Policies", title: "Vigil mechanism (whistle-blower) policy", short: "A policy for directors and employees to report concerns safely.",
    basis: "Section 177(9) and (10); Rule 7, Companies (Meetings of Board and its Powers) Rules, 2014",
    fields: POL_FIELDS.concat([{ id: "chair", label: "Chairperson of the Audit Committee" }]),
    notes: ["The mechanism is required for listed companies, companies that accept deposits from the public, and companies with borrowings from banks and financial institutions above Rs. 50 crore. Check whether you need it.", "Place the policy on the company's website if there is one, and report on it in the Board's report."],
    build(v) {
      const o = polHead(v, "VIGIL MECHANISM POLICY", "Under Section 177(9) of the Companies Act, 2013"); let n = 1;
      clause(o, n++, "Purpose", `${CO(v)} wants its directors and employees to be able to report genuine concerns about unethical behaviour, actual or suspected fraud, or a breach of its code of conduct, without fear. This policy gives them an adequate way to do so.`);
      clause(o, n++, "Who can use it", "Every director and employee of the Company.");
      clause(o, n++, "What can be reported", "Unethical behaviour, fraud or suspected fraud, a violation of law or of the Company's policies, misuse of the Company's assets, and a leak of confidential information.");
      clause(o, n++, "How to report", `A report may be made in writing, or by e-mail to ${va(v, "contact", "contact person")} at ${va(v, "email", "e-mail")}. In exceptional cases, the person can write directly to the Chairperson of the Audit Committee, ${va(v, "chair", "name")}.`);
      clause(o, n++, "Protection", "No person who makes a report in good faith will be dismissed, demoted, harassed or otherwise treated unfairly because of it. The identity of the person is kept confidential as far as the law and the inquiry allow.");
      clause(o, n++, "Inquiry", "Each report is looked into promptly and fairly. The person may be asked for more details. The Audit Committee reviews the reports and the action taken.");
      clause(o, n++, "Misuse", "A report found to be false and made with a bad intention may lead to disciplinary action.");
      clause(o, n++, "Records and review", "A record of all reports and the action taken is kept as the law requires. The Board may review and amend this policy from time to time.");
      return o;
    },
  },
  {
    id: "nrcpol", group: "Policies", title: "Nomination and remuneration policy", short: "Policy on director selection, independence and remuneration.",
    basis: "Section 178(2) to (4), Companies Act, 2013; Rule 6, Companies (Meetings of Board and its Powers) Rules, 2014",
    fields: POL_FIELDS.slice(0, 2),
    notes: ["The committee is required for listed companies and for public companies above the limits in Rule 6. Check that before adopting this policy.", "The committee must have three or more non-executive directors, at least half of them independent, with a non-executive chairperson."],
    build(v) {
      const o = polHead(v, "NOMINATION AND REMUNERATION POLICY", "Under Section 178 of the Companies Act, 2013"); let n = 1;
      clause(o, n++, "Purpose", `This policy sets out how ${CO(v)} chooses its directors, key managerial personnel and senior management, and how it fixes their pay.`);
      clause(o, n++, "The Committee", "The Nomination and Remuneration Committee has three or more non-executive directors, at least half of whom are independent. The chairperson is an independent director.");
      clause(o, n++, "Qualifications and positive attributes", "A director should have the integrity, knowledge, experience and time the role needs, should not be disqualified under the Act, and should hold no more directorships than the law allows.");
      clause(o, n++, "Independent directors", "An independent director must meet the criteria in Section 149(6). The Committee checks them at appointment and every year.");
      clause(o, n++, "Remuneration of directors", "Remuneration of a managing or whole-time director follows Sections 196 and 197 and Schedule V. A non-executive director may be paid sitting fees and, where the Act allows, commission, within the limits of Section 197 and the Rules.");
      clause(o, n++, "Key managerial personnel and senior management", "Their pay has a fixed part and, where suitable, a performance-linked part. It reflects the market and the person's role.");
      clause(o, n++, "Evaluation", "The Committee sets the way in which the Board, its committees and each director are evaluated every year.");
      clause(o, n++, "Removal", "A director or senior manager may be removed for reasons the Act and the terms of appointment allow, with the reasons recorded.");
      clause(o, n++, "Review", "The Committee reviews this policy as needed and recommends changes to the Board.");
      return o;
    },
  },
  {
    id: "csrpol", group: "Policies", title: "Corporate social responsibility policy", short: "A CSR policy covering the spend, the committee and the projects.",
    basis: "Section 135 and Schedule VII, Companies Act, 2013; Companies (Corporate Social Responsibility Policy) Rules, 2014",
    fields: POL_FIELDS.slice(0, 2).concat([{ id: "areas", label: "Areas the company will support", type: "textarea", ph: "Education, health care, environment" }]),
    notes: ["Section 135 applies when, in the immediately preceding year, the company had a net worth of Rs. 500 crore or more, or turnover of Rs. 1,000 crore or more, or net profit of Rs. 5 crore or more.", "The company must spend at least 2% of the average net profit of the last three years. Unspent amounts go to a separate account or fund within the time in the Rules.", "A company whose CSR spend is Rs. 50 lakh or less need not have a CSR Committee. Check the current rule."],
    build(v) {
      const o = polHead(v, "CORPORATE SOCIAL RESPONSIBILITY POLICY", "Under Section 135 of the Companies Act, 2013"); let n = 1;
      clause(o, n++, "Purpose", `${CO(v)} wants to contribute to the social and economic development of the communities in which it works, through the activities in Schedule VII to the Act.`);
      clause(o, n++, "Areas of focus", va(v, "areas", "areas") + ", and any other activity in Schedule VII that the Board approves.");
      clause(o, n++, "CSR Committee", "The Board forms a CSR Committee as the Act requires. The Committee recommends the policy, the annual action plan and the amount of spend, and watches over the projects.");
      clause(o, n++, "Amount to be spent", "The Company spends, in every financial year, at least two per cent of the average net profits of the preceding three financial years, worked out as the Act and the Rules require.");
      clause(o, n++, "Projects", "Projects may be carried out directly or through a registered trust, society or company that meets the conditions in the Rules. Each project has a plan, a budget, a time line and a way of checking the result.");
      clause(o, n++, "Unspent amount", "An unspent amount for an ongoing project is moved to a separate Unspent CSR Account within thirty days of the end of the year and spent within three years. Any other unspent amount is moved to a fund in Schedule VII within six months of the end of the year.");
      clause(o, n++, "Reporting", "The Board's report carries the annual report on CSR in the format the Rules give. The Company files the CSR form with the Registrar as the Rules require.");
      return o;
    },
  },
  {
    id: "posh", group: "Policies", title: "Policy on prevention of sexual harassment (POSH)", short: "A policy and Internal Committee set-up under the 2013 Act.",
    basis: "Sexual Harassment of Women at Workplace (Prevention, Prohibition and Redressal) Act, 2013",
    fields: POL_FIELDS.slice(0, 2).concat([{ id: "pres", label: "Presiding Officer of the Internal Committee (a senior woman employee)" }, { id: "mem", label: "Other members (one per line)", type: "textarea" }, { id: "ext", label: "External member" }]),
    notes: ["Every employer with ten or more employees must form an Internal Committee at each office. Its presiding officer must be a senior woman employee, at least half its members must be women, and one member must be from outside the company.", "Hold awareness sessions, display the notice, and report the number of complaints in the Board's report and to the District Officer each year."],
    build(v) {
      const o = polHead(v, "POLICY ON PREVENTION OF SEXUAL HARASSMENT AT THE WORKPLACE", "Under the Sexual Harassment of Women at Workplace (Prevention, Prohibition and Redressal) Act, 2013"); let n = 1;
      clause(o, n++, "Commitment", `${CO(v)} gives every woman who works with it, whether employed, contract, trainee or visiting, a workplace free from sexual harassment.`);
      clause(o, n++, "What counts as sexual harassment", "Unwelcome physical contact, a demand or request for sexual favours, sexually coloured remarks, showing pornography, or any other unwelcome physical, verbal or non-verbal conduct of a sexual nature.");
      clause(o, n++, "Internal Committee", `The Company has an Internal Committee headed by ${va(v, "pres", "presiding officer")}, a senior woman employee. Its other members are: ${LN("mem")(v).join("; ") || "[members]"}. The external member is ${va(v, "ext", "external member")}. At least half of the members are women.`);
      clause(o, n++, "Making a complaint", `An aggrieved woman may give a written complaint to the Internal Committee within three months of the incident, or of the last incident in a series. The Committee may extend that by another three months for good reasons. If she cannot write, the Committee will help her. Contact: ${va(v, "contact", "name")}, ${va(v, "email", "e-mail")}.`);
      clause(o, n++, "Inquiry", "The Committee gives notice to both parties, hears them and any witnesses, keeps the matter confidential, and finishes the inquiry within ninety days. It gives its report within ten days of finishing. The employer acts on the report within sixty days.");
      clause(o, n++, "Conciliation", "Before the inquiry, and if the woman asks, the Committee may try to settle the matter through conciliation. No money may be taken as a settlement.");
      clause(o, n++, "Action", "If the complaint is proved, the Company takes action as its service rules provide, and may deduct compensation from the respondent's pay. A false or malicious complaint may also lead to action.");
      clause(o, n++, "Awareness and reporting", "The Company holds awareness programmes, displays this policy, and reports the number of complaints and their disposal in the Board's report and in the annual report to the District Officer.");
      return o;
    },
  },
);

/* packs: capital, name, objects ------------------------------------------------- */
function packEGM(v, bizTitle, kind, extra) {
  const o = [];
  o.push(...header(v), T("NOTICE OF EXTRAORDINARY GENERAL MEETING"), R("Date: " + vd(v, "ndate", "date of notice")));
  o.push(P(`**NOTICE** is hereby given that an Extraordinary General Meeting of the members of **${CO(v)}** will be held on **${vd(v, "egmdate", "date")}** at **${va(v, "egmtime", "time")}** at ${va(v, "venue", "venue")}, to transact the following business:`));
  o.push(P("**SPECIAL BUSINESS**"), P(`**1.** To consider and, if thought fit, to pass the following resolution as ${kind === "Special" ? "a Special" : "an Ordinary"} Resolution: ${bizTitle}`));
  o.push(...extra, S(), L("By order of the Board"), R("For **" + CO(v) + "**"), S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")));
  o.push(S(), P("**Notes:** A member entitled to attend and vote may appoint a proxy, who need not be a member. The proxy form must reach the registered office at least 48 hours before the meeting. The explanatory statement under Section 102 is annexed."));
  o.push(S(), P("**EXPLANATORY STATEMENT UNDER SECTION 102**"), P(GAP));
  return o;
}
const PACK_COMMON = [
  { id: "mdate", label: "Date of Board meeting", type: "date" }, { id: "ndate", label: "Date of EGM notice", type: "date" }, { id: "egmdate", label: "Date of EGM", type: "date" }, { id: "egmtime", label: "Time of EGM", ph: "11:00 a.m." }, { id: "venue", label: "Venue of EGM", type: "textarea" },
  { id: "sname", label: "Name of the signing director or officer" }, { id: "sdes", label: "Designation" },
];
GENS.push(
  {
    id: "capitalpack", group: "Packs", title: "Increase in authorised capital: pack", short: "Board resolution, EGM notice, shareholders' resolution and the altered Clause V.",
    basis: "Sections 61 and 64, Companies Act, 2013; Form SH-7",
    fields: PACK_COMMON.concat([{ id: "oldcap", label: "Present authorised capital (Rs.)", type: "number" }, { id: "newcap", label: "New authorised capital (Rs.)", type: "number" }, { id: "fv", label: "Face value per share (Rs.)", type: "number" }]),
    notes: ["The Articles must authorise the increase. If they do not, alter them first by a special resolution.", "File Form SH-7 within 30 days of the resolution, with the altered Memorandum. Check whether Form MGT-14 is also needed for your company.", "Stamp duty on the altered Memorandum differs by State."],
    build(v) {
      const nc = String(v.newcap || "").replace(/[^0-9]/g, ""), fv = String(v.fv || "").replace(/[^0-9]/g, ""), cnt = nc && fv ? Math.floor(+nc / +fv) : "";
      const o = [...header(v), T("1. RESOLUTION PASSED AT THE MEETING OF THE BOARD OF DIRECTORS HELD ON " + vd(v, "mdate", "date").toUpperCase())];
      o.push(RES(`subject to the approval of the members, the authorised share capital of the Company be increased from ${vmoney(v, "oldcap", "present capital")} to ${vmoney(v, "newcap", "new capital")}, divided into ${cnt ? vnum({ x: cnt }, "x") : "[number]"} equity shares of Rs. ${va(v, "fv", "face value")} each, by the creation of additional shares ranking equally with the existing shares, and Clause V of the Memorandum of Association and the relevant Article be altered accordingly.`));
      o.push(RESF(`an Extraordinary General Meeting be held on ${vd(v, "egmdate", "date")} to seek the approval of the members, and the notice of that meeting be approved and issued.`)); o.push(...signoff(v, va(v, "sname", "name"), va(v, "sdes", "designation")), PB());
      o.push(...packEGM(v, `**RESOLVED THAT** pursuant to Sections 61 and 64 of the Companies Act, 2013, the Articles of Association and the Memorandum of Association, the authorised share capital of the Company be increased from ${vmoney(v, "oldcap", "present capital")} to ${vmoney(v, "newcap", "new capital")} by the creation of additional equity shares of Rs. ${va(v, "fv", "face value")} each, and that Clause V of the Memorandum of Association be substituted as set out below.`, "Ordinary", [P(`**Clause V.** The authorised share capital of the Company is ${vmoney(v, "newcap", "new capital")} divided into ${cnt ? vnum({ x: cnt }, "x") : "[number]"} equity shares of Rs. ${va(v, "fv", "face value")} each, with power to increase or reduce the capital and to divide the shares into classes, as the Act and the Articles allow.`)]));
      o.push(S(), P("**After the meeting:** file Form SH-7 with the Registrar of Companies within thirty days, attaching the notice, the resolution and the altered Memorandum."));
      return o;
    },
  },
  {
    id: "namepack", group: "Packs", title: "Change of company name: pack", short: "Board resolution, EGM notice and special resolution, and the altered Clause I.",
    basis: "Sections 4, 13 and 117, Companies Act, 2013; Rule 29, Companies (Incorporation) Rules, 2014; Form INC-24",
    fields: PACK_COMMON.concat([{ id: "oldname", label: "Present name" }, { id: "newname", label: "New name" }]),
    notes: ["Reserve the new name first with the Registrar. The name change takes effect only after the Registrar approves it and issues a fresh certificate of incorporation.", "File Form INC-24 within 30 days of the special resolution, and file Form MGT-14 for the special resolution as required.", "Alter the Memorandum and the Articles and inform the bank, the tax authorities and others."],
    build(v) {
      const on = va(v, "oldname", "present name"), nn = va(v, "newname", "new name");
      const o = [...header(v), T("1. RESOLUTION PASSED AT THE MEETING OF THE BOARD OF DIRECTORS HELD ON " + vd(v, "mdate", "date").toUpperCase())];
      o.push(RES(`subject to the approval of the members and the Central Government (through the Registrar of Companies), the name of the Company be changed from ${on} to ${nn}, the name having been made available by the Registrar of Companies.`));
      o.push(RESF(`an Extraordinary General Meeting be held on ${vd(v, "egmdate", "date")} to approve the change, the notice be approved, and Clause I of the Memorandum and the relevant Article of the Articles of Association be altered accordingly.`)); o.push(...signoff(v, va(v, "sname", "name"), va(v, "sdes", "designation")), PB());
      o.push(...packEGM(v, `**RESOLVED THAT** pursuant to Section 13 and other applicable provisions of the Companies Act, 2013, and subject to the approval of the Central Government, the name of the Company be changed from ${on} to ${nn}, and that Clause I of the Memorandum of Association and the Articles of Association be altered by replacing the old name with the new name wherever it appears.`, "Special", [P(`**Clause I.** The name of the Company is ${nn}.`)]));
      o.push(S(), P("**After the meeting:** file Form INC-24 with the Registrar of Companies within thirty days of the special resolution, with the altered Memorandum and Articles and the other documents the Rules require, and file Form MGT-14."));
      return o;
    },
  },
  {
    id: "objectspack", group: "Packs", title: "Change of objects clause: pack", short: "Board resolution, EGM notice and special resolution, and the altered Clause III.",
    basis: "Sections 13(1) and 117, Companies Act, 2013; Form MGT-14",
    fields: PACK_COMMON.concat([{ id: "objs", label: "Objects to be added or substituted", type: "textarea" }]),
    notes: ["A change to the objects clause needs a special resolution of the members. Where the company has raised money by a prospectus and not used it, extra conditions apply (Section 13(8)).", "File Form MGT-14 within 30 days with the altered Memorandum. Other approvals may be needed if the new business is regulated."],
    build(v) {
      const o = [...header(v), T("1. RESOLUTION PASSED AT THE MEETING OF THE BOARD OF DIRECTORS HELD ON " + vd(v, "mdate", "date").toUpperCase())];
      o.push(RES("subject to the approval of the members by a special resolution, the objects clause (Clause III) of the Memorandum of Association of the Company be altered as set out in the draft placed before the Board."));
      o.push(RESF(`an Extraordinary General Meeting be held on ${vd(v, "egmdate", "date")}, and the notice of that meeting be approved and issued.`)); o.push(...signoff(v, va(v, "sname", "name"), va(v, "sdes", "designation")), PB());
      o.push(...packEGM(v, `**RESOLVED THAT** pursuant to Section 13 and other applicable provisions of the Companies Act, 2013, Clause III of the Memorandum of Association of the Company be altered by adding or substituting the objects set out below.`, "Special", [P("**Clause III. Objects to be added or substituted:**"), P(va(v, "objs", "objects"))]));
      o.push(S(), P("**After the meeting:** file Form MGT-14 within thirty days, attaching the altered Memorandum, the notice and the explanatory statement."));
      return o;
    },
  },
);


/* ---------- blocks -> html (preview) and docx ---------- */
function inlineHtml(t) { return esc(t).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/\n/g, "<br>"); }
function blocksHtml(bs) {
  return bs.map(b => {
    if (b.k === "S") return '<p class="sp"></p>'; if (b.k === "PB") return '<hr class="pb">';
    const c = { T: "t", H: "h", P: "p", L: "l", C: "c", R: "r", I: "i" }[b.k]; return `<p class="${c}"${b.k === "I" ? ` style="margin-left:${2 * (b.n || 1)}rem"` : ""}>${inlineHtml(b.t)}</p>`;
  }).join("");
}
function blocksText(bs) { return bs.map(b => b.k === "S" || b.k === "PB" ? "" : String(b.t).replace(/\*\*/g, "")).join("\n").replace(/\n{3,}/g, "\n\n"); }

const X = s => String(s).replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
function runs(t, base) {
  const parts = String(t).split(/(\*\*.+?\*\*)/g).filter(x => x !== ""); let o = "";
  parts.forEach(p => { const bold = /^\*\*.+\*\*$/.test(p); const txt = bold ? p.slice(2, -2) : p; o += `<w:r><w:rPr>${bold || base.b ? "<w:b/>" : ""}${base.u ? '<w:u w:val="single"/>' : ""}${base.sz ? `<w:sz w:val="${base.sz}"/>` : ""}</w:rPr><w:t xml:space="preserve">${X(txt)}</w:t></w:r>`; });
  return o;
}
function docxBody(bs) {
  return bs.map(b => {
    if (b.k === "PB") return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    if (b.k === "S") return '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>';
    const jc = { T: "center", H: "center", C: "center", P: "both", L: "left", R: "right", I: "both" }[b.k];
    const base = b.k === "T" ? { b: 1, u: 0, sz: 26 } : b.k === "H" ? { b: 1, sz: 26 } : {};
    const ind = b.k === "I" ? `<w:ind w:left="${720 * (b.n || 1)}"/>` : "";
    return `<w:p><w:pPr><w:spacing w:after="140" w:line="300" w:lineRule="auto"/>${ind}<w:jc w:val="${jc}"/></w:pPr>${runs(b.t, base)}</w:p>`;
  }).join("");
}
/* tiny zip writer (stored, no compression) */
const CRC = (() => { const t = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(u) { let c = 0xFFFFFFFF; for (let i = 0; i < u.length; i++) c = CRC[(c ^ u[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function zip(files) {
  const enc = new TextEncoder(), chunks = [], cd = []; let off = 0;
  const u16 = n => [n & 255, (n >> 8) & 255], u32 = n => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
  files.forEach(f => {
    const name = enc.encode(f.n), data = enc.encode(f.d), crc = crc32(data);
    const lh = [0x50, 0x4b, 3, 4, ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length), ...u16(0)];
    chunks.push(new Uint8Array(lh), name, data);
    cd.push(new Uint8Array([0x50, 0x4b, 1, 2, ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(off)]), name);
    off += lh.length + name.length + data.length;
  });
  let cs = 0; cd.forEach(c => cs += c.length);
  const end = new Uint8Array([0x50, 0x4b, 5, 6, 0, 0, 0, 0, ...u16(files.length), ...u16(files.length), ...u32(cs), ...u32(off), 0, 0]);
  return new Blob([...chunks, ...cd, end], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}
function makeDocx(bs) {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  return zip([
    { n: "[Content_Types].xml", d: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>' },
    { n: "_rels/.rels", d: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' },
    { n: "word/_rels/document.xml.rels", d: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
    { n: "word/styles.xml", d: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-IN"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>` },
    { n: "word/document.xml", d: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${docxBody(bs)}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1304" w:bottom="1440" w:left="1304" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>` },
  ]);
}

/* ---------- the page ---------- */
const S_ = { cur: null, vals: {} };
function loadCommon() { try { return JSON.parse(store.get("gen_company") || "{}"); } catch (e) { return {}; } }
function fieldHtml(f, v) {
  const val = v[f.id] != null ? v[f.id] : (f.def || ""); const id = "f_" + f.id; const req = f.req ? " *" : "";
  let el;
  if (f.type === "textarea") el = `<textarea id="${id}" data-f="${f.id}" rows="3" placeholder="${esc(f.ph || "")}">${esc(val)}</textarea>`;
  else if (f.type === "select") el = `<select id="${id}" data-f="${f.id}">${f.opts.map(o => `<option value="${esc(o[0])}"${String(val || f.opts[0][0]) === o[0] ? " selected" : ""}>${esc(o[1])}</option>`).join("")}</select>`;
  else el = `<input id="${id}" data-f="${f.id}" type="${f.type === "date" ? "date" : f.type === "number" ? "text" : "text"}"${f.type === "number" ? ' inputmode="decimal"' : ""} value="${esc(val)}" placeholder="${esc(f.ph || "")}">`;
  return `<div class="fld" data-w="${f.id}"><label for="${id}">${esc(f.label)}${req}</label>${el}</div>`;
}
function currentVals() { const v = {}; $$("[data-f]").forEach(e => v[e.dataset.f] = e.value); return v; }
function refresh() {
  const g = S_.cur; if (!g) return; const v = currentVals(); S_.vals[g.id] = v;
  const c = {}; COMMON.forEach(f => c[f.id] = v[f.id] || ""); store.set("gen_company", JSON.stringify(c));
  g.fields.forEach(f => { const w = $(`[data-w="${f.id}"]`); if (w) w.hidden = !!(f.show && !f.show(v)); });
  const bs = g.build(v); S_.blocks = bs; $("#doc").innerHTML = blocksHtml(bs);
}
function openGen(id) {
  const g = GENS.find(x => x.id === id); if (!g) return listGens(); openGenObj(g);
}
function openGenObj(g) {
  S_.cur = g;
  const saved = S_.vals[g.id] || {}, com = loadCommon(), v = Object.assign({}, com, saved);
  $("#list").hidden = true; $("#work").hidden = false; document.title = g.title + " | Document generators";
  $("#work").innerHTML = `<p class="small"><a href="#" id="back">All generators</a></p><h1>${esc(g.title)}</h1><p class="lede">${esc(g.short)}</p><p class="basis">Based on: ${esc(g.basis)}</p>
  <div class="cols"><form class="form" id="form" autocomplete="off"><h2>Company</h2>${COMMON.map(f => fieldHtml(f, v)).join("")}<h2>Details</h2>${g.fields.map(f => fieldHtml(f, v)).join("")}</form>
  <div class="prev"><div class="actions"><button class="btn pri" id="dl">Download Word file</button><button class="btn" id="cp">Copy text</button><button class="btn" id="pr">Print</button><span id="msg" class="msg" aria-live="polite"></span></div>
  <div class="paper" id="doc"></div>
  <div class="check"><h2>Check before you use it</h2><ul>${g.notes.map(n => `<li>${esc(n)}</li>`).join("")}</ul><p class="small">This is a starting draft made from the Act and the Rules. Have it reviewed by a Company Secretary or a lawyer before it is signed or filed.</p></div></div></div>`;
  $("#form").addEventListener("input", refresh); $("#form").addEventListener("change", refresh);
  $("#back").onclick = e => { e.preventDefault(); location.hash = ""; };
  $("#dl").onclick = () => { const b = makeDocx(S_.blocks), a = document.createElement("a"); a.href = URL.createObjectURL(b); a.download = (g.id + "-" + ((currentVals().company || "company").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-|-$/g, "")) + ".docx"); document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500); $("#msg").textContent = "Downloaded."; };
  $("#cp").onclick = async () => { const t = blocksText(S_.blocks); try { await navigator.clipboard.writeText(t); $("#msg").textContent = "Copied."; } catch (e) { const ta = document.createElement("textarea"); ta.value = t; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); ta.remove(); $("#msg").textContent = "Copied."; } };
  $("#pr").onclick = () => window.print();
  refresh(); window.scrollTo(0, 0);
}
function listGens() {
  S_.cur = null; document.title = "Document generators | Company Law Knowledge Lab"; $("#work").hidden = true; $("#list").hidden = false;
  const groups = []; GENS.forEach(g => { if (groups.indexOf(g.group) < 0) groups.push(g.group); });
  $("#list").innerHTML = `<h1>Document generators</h1><p class="lede">Fill in the details once and download an editable Word file. Everything runs in your browser: nothing you type is sent anywhere. Your company details are remembered on this device only.</p>
  ${groups.map(gr => `<h2>${esc(gr)}</h2><div class="cards">${GENS.filter(g => g.group === gr).map(g => `<a class="card" href="#${g.id}"><b>${esc(g.title)}</b><span>${esc(g.short)}</span><i>${esc(g.basis)}</i></a>`).join("")}</div>`).join("")}
  <h2>Model resolutions from the book</h2><div id="mres"><p class="small">Loading…</p></div>`;
  mresInit();
}
/* ---------- model resolutions (from the library's data) ---------- */
const MR = { idx: null, part: "A", q: "", lim: 40 };
async function mresInit() {
  const box = $("#mres"); if (!box) return;
  try { if (!MR.idx) { const r = await fetch("../data/res/index.json"); if (!r.ok) throw 0; MR.idx = await r.json(); } }
  catch (e) { box.innerHTML = '<p class="small">The model resolutions are not available yet. Run the "Build data and website" workflow once, and they will appear here.</p>'; return; }
  const parts = MR.idx.parts, cnt = p => MR.idx.items.filter(i => i.p === p).length;
  box.innerHTML = `<p class="lede">All ${MR.idx.items.length} model resolutions from the book kept in the library. Pick one, fill in the blanks, and download it as a Word file with your company's name on it.</p>
  <div class="mtabs">${Object.keys(parts).map(p => `<button class="btn${MR.part === p ? " pri" : ""}" data-p="${p}">${esc(parts[p])} (${cnt(p)})</button>`).join(" ")}</div>
  <div class="fld"><label for="mq">Search the resolutions</label><input id="mq" type="text" placeholder="for example: dividend, auditor, borrowing" value="${esc(MR.q)}"></div><div id="mlist"></div>`;
  $$(".mtabs button", box).forEach(b => b.onclick = () => { MR.part = b.dataset.p; MR.lim = 40; mresInit(); });
  $("#mq").oninput = e => { MR.q = e.target.value; MR.lim = 40; mresList(); }; mresList();
}
function mresList() {
  const q = MR.q.toLowerCase().split(/\s+/).filter(Boolean);
  const it = MR.idx.items.filter(i => i.p === MR.part && q.every(w => (i.t + " " + i.m + " " + i.g + " " + i.s.join(" ") + " " + i.f.join(" ")).toLowerCase().indexOf(w) >= 0));
  $("#mlist").innerHTML = `<div class="rlist">${it.slice(0, MR.lim).map(i => `<a class="rrow" href="#res/${esc(i.id)}"><b>${esc(i.t)}</b><span>${esc(i.m)}${i.g ? " · " + esc(i.g) : ""}${i.s.length ? " · Section " + esc(i.s.slice(0, 4).join(", ")) : ""}</span></a>`).join("") || '<p class="small">Nothing matches that search.</p>'}</div>${it.length > MR.lim ? `<p><button class="btn" id="mmore">Show more (${it.length - MR.lim} left)</button></p>` : ""}`;
  const m = $("#mmore"); if (m) m.onclick = () => { MR.lim += 60; mresList(); };
}
async function openRes(id) {
  let d; try { const r = await fetch("../data/res/" + encodeURIComponent(id) + ".json"); if (!r.ok) throw 0; d = await r.json(); } catch (e) { location.hash = ""; return; }
  const all = [].concat(d.proposal, d.resolution, d.minutes), fields = [], seen = [];
  all.forEach(t => { let m; const re = /_{3,}/g; while ((m = re.exec(t))) { const before = t.slice(0, m.index).replace(/_+/g, " ").trim().split(/\s+/).slice(-6).join(" "); fields.push({ id: "b" + fields.length, label: `Blank ${fields.length + 1}: after "${before || "start"}"`, ph: "" }); } });
  const fill = (t, st, v) => t.replace(/_{3,}/g, (m, off, str) => { const id = "b" + st.n++; const x = (v[id] || "").trim(); if (x) return x; if (/^\s*(Private\s+)?(Limited|Ltd)\b/i.test(str.slice(off + m.length)) && (v.company || "").trim()) return v.company.trim().replace(/\s*(Private\s+)?(Limited|Ltd\.?)$/i, ""); return "________"; });
  const g = {
    id: "res-" + id, group: "", title: d.t, short: [d.m, d.g].filter(Boolean).join(" · "), basis: "Book of model resolutions kept in the library" + (d.s.length ? "; Section " + d.s.slice(0, 6).join(", ") : "") + (d.f.length ? "; Forms " + d.f.join(", ") : ""),
    fields: [{ id: "mdate", label: "Date of the meeting", type: "date" }, { id: "place", label: "Place of the meeting" }, { id: "sname", label: "Name of the signing director or Chairperson" }].concat(fields),
    notes: ["The wording comes from the book of model resolutions kept in the library (2016 edition). Laws and forms change, so check every resolution against the current Act and Rules before using it.", "Fill in the blanks. Where a blank comes before 'Limited', the company name is used if you leave it empty.", "Remove any paragraph that does not fit your company's facts. The 2016 text may mention rules that have since changed, such as Dividend Distribution Tax, which no longer applies."],
    build(v) {
      const st = { n: 0 }, out = header(v); out.push(T(d.t.toUpperCase()));
      out.push(L([d.m, v.mdate ? "Meeting held on " + fdate(v.mdate) : "", v.place ? "at " + v.place : ""].filter(Boolean).join(" · ")), S());
      if (d.proposal.length) { out.push(P("**Proposal**")); d.proposal.forEach(t => out.push(P(fill(t, st, v)))); }
      if (d.resolution.length) { out.push(P("**Proposed resolution**")); d.resolution.forEach(t => out.push(P(fill(t, st, v)))); }
      if (d.minutes.length) { out.push(P("**Minutes: gist of the discussion**")); d.minutes.forEach(t => out.push(P(fill(t, st, v)))); }
      return out.concat(signoff(v, va(v, "sname", "name"), "Director / Chairperson"));
    },
  };
  openGenObj(g);
  const bk = $("#back"); if (bk) bk.onclick = e => { e.preventDefault(); location.hash = ""; };
}
function route() { const id = location.hash.replace(/^#/, ""); if (id.indexOf("res/") === 0) openRes(id.slice(4)); else if (id) openGen(id); else listGens(); }
window.addEventListener("hashchange", route);
(function theme() { const t = store.get("theme") || "auto"; document.documentElement.dataset.theme = t; $$(".sw").forEach(b => { b.setAttribute("aria-pressed", String(b.dataset.t === t)); b.onclick = () => { store.set("theme", b.dataset.t); theme(); }; }); })();
route();
"""

FONTS = '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=Source+Serif+4:ital,opsz,wght@0,8..60,400;0,8..60,600;1,8..60,400&family=Zilla+Slab:wght@400;500;600;700&display=swap" rel="stylesheet">'


def shell(title, desc, active, body, script=""):
    nav = [("../", "Home", ""), ("../#/act", "Companies Act", ""), ("../#/forms", "Forms", ""), ("../#/processes", "Processes", ""),
           ("../generators/", "Document generators", "gen"), ("../excel/", "Excel utilities", "xl"), ("../drafts/", "Drafts", "")]
    cur = ' aria-current="page"'
    links = "".join('<a href="%s"%s>%s</a>' % (h, cur if (k and k == active) else "", t) for h, t, k in nav)
    return f"""<!doctype html>
<html lang="en" data-theme="auto"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title><meta name="description" content="{desc}">{FONTS}<style>{CSS}</style></head>
<body><header class="bar"><div class="bar-in"><a class="brand" href="../">Company Law Knowledge Lab</a><nav aria-label="Sections">{links}</nav>
<div class="themes" role="group" aria-label="Colour theme"><button class="sw" data-t="light" aria-label="Light" aria-pressed="false"></button><button class="sw" data-t="sepia" aria-label="Cream" aria-pressed="false"></button><button class="sw" data-t="dark" aria-label="Dark" aria-pressed="false"></button></div></div></header>
<main>{body}</main>
<footer>For study only. These tools make starting drafts and working files from the Act and the Rules. They are not legal advice: check every output against the current law before you rely on it.</footer>
<script>{script}</script></body></html>"""


def build_generators():
    out = SITE / "generators"
    out.mkdir(parents=True, exist_ok=True)
    body = '<div id="list"></div><div id="work" hidden></div>'
    (out / "index.html").write_text(shell("Document generators | Company Law Knowledge Lab", "Fill in a form and download an editable Word file: Board resolutions, notices, minutes, letter of authority and share certificate.", "gen", body, GEN_JS), encoding="utf-8")
    print("Document generators page written")


# ---------------------------------------------------------------- Excel
def build_excel():
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    NAVY, GOLD, LIGHT, INPUT = "0F2747", "C8A13A", "E9EEF5", "FFF8DC"
    thin = Side(style="thin", color="C9D3E1")
    BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
    DATEF = "dd-mmm-yyyy"
    out = SITE / "excel"
    out.mkdir(parents=True, exist_ok=True)
    made = []
    GRP = [""]

    def head(ws, row, headers, widths):
        for i, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(row=row, column=i, value=h)
            c.font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            c.border = BOX
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[row].height = 46

    def title(ws, text, sub):
        ws["A1"] = text
        ws["A1"].font = Font(bold=True, size=15, color=NAVY)
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, size=10, color="4B5A70")

    def readme(wb, name, basis, lines):
        ws = wb.create_sheet("Read me", 0)
        ws.column_dimensions["A"].width = 120
        ws["A1"] = name
        ws["A1"].font = Font(bold=True, size=16, color=NAVY)
        ws["A2"] = "Based on: " + basis
        ws["A2"].font = Font(italic=True, color="4B5A70")
        r = 4
        for ln in lines + ["", "Cells with a light yellow fill are for you to fill in. Cells with a grey-blue fill are worked out by formulas: please don't type over them.",
                           "Check every date and rule against the current Act, Rules and circulars. Extensions and exemptions are not built into these files.",
                           "Free starting file from the Company Law Knowledge Lab. For study only; not legal advice."]:
            ws.cell(row=r, column=1, value=ln).alignment = Alignment(wrap_text=True, vertical="top")
            r += 1

    def style_rows(ws, first, last, ncols, formula_cols=(), date_cols=(), num_cols=(), wrap_cols=()):
        for r in range(first, last + 1):
            for c in range(1, ncols + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = BOX
                cell.alignment = Alignment(vertical="top", wrap_text=(c in wrap_cols))
                if c in formula_cols:
                    cell.fill = PatternFill("solid", fgColor=LIGHT)
                else:
                    cell.fill = PatternFill("solid", fgColor=INPUT)
                if c in date_cols:
                    cell.number_format = DATEF
                if c in num_cols:
                    cell.number_format = "#,##0.00"

    def inputs(ws, items, col_label=1, col_val=2, first_row=3):
        for i, (label, fmt) in enumerate(items):
            r = first_row + i
            ws.cell(row=r, column=col_label, value=label).font = Font(bold=True)
            c = ws.cell(row=r, column=col_val)
            c.fill = PatternFill("solid", fgColor=INPUT)
            c.border = BOX
            if fmt:
                c.number_format = fmt

    def save(wb, fn, name, blurb, basis, sheets):
        wb.save(out / fn)
        made.append({"file": fn, "name": name, "blurb": blurb, "basis": basis, "sheets": sheets, "group": GRP[0]})

    def status_rules(ws, rng, col, first):
        ref = f"{col}{first}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{ref}="Overdue"'], fill=PatternFill("solid", bgColor="F4B6AD", fgColor="F4B6AD"), font=Font(bold=True, color="8A1C0E")))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({ref},3)="Due"'], fill=PatternFill("solid", bgColor="FFE08A", fgColor="FFE08A")))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{ref}="Filed"'], fill=PatternFill("solid", bgColor="B7E1C4", fgColor="B7E1C4")))

    GRP[0] = "Filings, meetings and planners"
    # 1. Annual ROC compliance tracker -------------------------------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Annual tracker"
    title(ws, "Annual ROC compliance tracker", "Enter the financial year end and the AGM date. The due dates and the status work themselves out.")
    inputs(ws, [("Company name", None), ("CIN", None), ("Financial year ends on", DATEF), ("AGM date (planned or held)", DATEF)], 1, 2, 3)
    ws["A7"] = "Today"
    ws["A7"].font = Font(bold=True)
    ws["B7"] = "=TODAY()"
    ws["B7"].number_format = DATEF
    hdr = ["Sr", "Compliance", "Form", "Law", "How the due date is worked out", "Due date", "Filed on", "Status", "SRN or reference", "Remarks"]
    head(ws, 9, hdr, [5, 40, 12, 24, 40, 14, 14, 18, 20, 30])
    rows = [
        ("Hold the Annual General Meeting", "-", "Section 96", "Within 6 months of the financial year end", '=IF($B$5="","",EDATE($B$5,6))'),
        ("File financial statements", "AOC-4", "Section 137", "30 days from the AGM", '=IF($B$6="","",$B$6+30)'),
        ("File annual return", "MGT-7 / MGT-7A", "Section 92", "60 days from the AGM", '=IF($B$6="","",$B$6+60)'),
        ("Appointment of auditor", "ADT-1", "Section 139", "15 days from the AGM (when an auditor is appointed)", '=IF($B$6="","",$B$6+15)'),
        ("File resolutions passed at the AGM", "MGT-14", "Section 117", "30 days from the AGM (for special resolutions)", '=IF($B$6="","",$B$6+30)'),
        ("Return of deposits and outstanding receipts", "DPT-3", "Rules 16 and 16A, Deposits Rules", "30 June (for a March year end)", '=IF($B$5="","",DATE(YEAR($B$5),6,30))'),
        ("Return on outstanding dues to small enterprises, April to September", "MSME-1", "Section 405", "31 October of the year before the year end", '=IF($B$5="","",DATE(YEAR($B$5)-1,10,31))'),
        ("Return on outstanding dues to small enterprises, October to March", "MSME-1", "Section 405", "30 April after the year end", '=IF($B$5="","",DATE(YEAR($B$5),4,30))'),
        ("KYC of every director with a DIN", "DIR-3 KYC", "Rule 12A, Appointment Rules", "30 September each year", '=IF($B$5="","",DATE(YEAR($B$5),9,30))'),
    ]
    for i, (a, f, law, how, fm) in enumerate(rows):
        r = 10 + i
        for c, v in enumerate([i + 1, a, f, law, how, fm], 1):
            ws.cell(row=r, column=c, value=v)
        ws.cell(row=r, column=8, value=f'=IF(F{r}="","",IF(G{r}<>"","Filed",IF(F{r}<$B$7,"Overdue",IF(F{r}-$B$7<=15,"Due in "&(F{r}-$B$7)&" days","Upcoming"))))')
    last = 10 + len(rows) + 5
    style_rows(ws, 10, last, 10, formula_cols=(6, 8), date_cols=(6, 7), wrap_cols=(2, 4, 5, 10))
    for r in range(10, 10 + len(rows)):
        for c in (1, 2, 3, 4, 5):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFFFFF")
    status_rules(ws, f"H10:H{last}", "H", 10)
    ws.freeze_panes = "A10"
    ws.cell(row=last + 2, column=2, value="Dates are for a company with a March year end. Extensions announced by the Ministry are not included.").font = Font(italic=True, color="4B5A70")
    w2 = wb.create_sheet("Event-based filings")
    title(w2, "Filings that run from an event", "Type the date of the event. The due date and the status work themselves out.")
    head(w2, 4, ["Sr", "Event", "Form", "Law", "Days allowed", "Date of event", "Due date", "Filed on", "Status", "SRN or reference", "Remarks"], [5, 44, 14, 26, 10, 14, 14, 14, 18, 20, 30])
    w2["A3"] = "Today"
    w2["A3"].font = Font(bold=True)
    w2["B3"] = "=TODAY()"
    w2["B3"].number_format = DATEF
    ev = [("Appointment or change of a director or key managerial personnel", "DIR-12", "Section 170(2)", 30), ("Special resolution or other resolution passed", "MGT-14", "Section 117", 30),
          ("Allotment of shares (30 days; 15 days for a private placement)", "PAS-3", "Section 39(4); Rule 14", 30), ("Creation or change of a charge", "CHG-1 / CHG-9", "Section 77", 30),
          ("Satisfaction of a charge", "CHG-4", "Section 82", 30), ("Increase in authorised share capital", "SH-7", "Section 64", 30)]
    for i in range(14):
        r = 5 + i
        if i < len(ev):
            for c, v in enumerate([i + 1, ev[i][0], ev[i][1], ev[i][2], ev[i][3]], 1):
                w2.cell(row=r, column=c, value=v)
        else:
            w2.cell(row=r, column=1, value=i + 1)
        w2.cell(row=r, column=7, value=f'=IF(OR(F{r}="",E{r}=""),"",F{r}+E{r})')
        w2.cell(row=r, column=9, value=f'=IF(G{r}="","",IF(H{r}<>"","Filed",IF(G{r}<$B$3,"Overdue",IF(G{r}-$B$3<=7,"Due in "&(G{r}-$B$3)&" days","Upcoming"))))')
    style_rows(w2, 5, 18, 11, formula_cols=(7, 9), date_cols=(6, 7, 8), wrap_cols=(2, 4, 11))
    status_rules(w2, "I5:I18", "I", 5)
    w2.freeze_panes = "A5"
    readme(wb, "Annual ROC compliance tracker", "Companies Act, 2013 and the Rules named against each line",
           ["1. On the 'Annual tracker' sheet, fill in the financial year end and the AGM date.", "2. The due date of each annual filing appears at once. Type the date you filed in the 'Filed on' column and the status turns green.",
            "3. The 'Event-based filings' sheet is for filings that run from an event, such as an allotment or the appointment of a director.", "4. 'Overdue' shows in red. 'Due in n days' shows in amber when 15 days or fewer are left (7 days for event-based filings)."])
    save(wb, "annual-roc-compliance-tracker.xlsx", "Annual ROC compliance tracker", "Due dates for the AGM, AOC-4, MGT-7, ADT-1, DPT-3, MSME-1 and DIR-3 KYC, with live status. Plus event-based filings such as DIR-12, MGT-14 and PAS-3.",
         "Sections 92, 96, 117, 137, 139, 170", ["Annual tracker", "Event-based filings"])

    # 2. Board and committee meeting tracker -------------------------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Meetings"
    title(ws, "Board and committee meeting tracker", "Checks notice, quorum, the gap between Board meetings and the dates for the minutes.")
    inputs(ws, [("Company name", None), ("Financial year", None), ("Total number of directors", "0")], 1, 2, 3)
    ws["D3"], ws["E3"] = "Board meetings held", '=COUNTIF($B$8:$B$207,"Board")'
    ws["D4"], ws["E4"] = "At least 4 are needed in a year", '=IF(E3>=4,"Met","Fewer than 4")'
    for c in ("D3", "D4"):
        ws[c].font = Font(bold=True)
    hdr = ["Sr", "Type of meeting", "Meeting date", "Notice sent on", "Days of notice", "Notice check (7 days)", "Directors present", "Quorum needed", "Quorum met?", "Days since the last Board meeting", "Gap check (120 days)",
           "Draft minutes to be sent by (15 days)", "Minutes to be signed by (30 days)", "Minutes signed on", "Remarks"]
    head(ws, 7, hdr, [5, 20, 14, 14, 10, 18, 11, 10, 11, 14, 18, 16, 16, 14, 28])
    for i in range(200):
        r = 8 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=5, value=f'=IF(OR(C{r}="",D{r}=""),"",C{r}-D{r})')
        ws.cell(row=r, column=6, value=f'=IF(E{r}="","",IF(E{r}>=7,"OK","Shorter than 7 days"))')
        ws.cell(row=r, column=8, value=f'=IF(OR($B$5="",C{r}=""),"",MAX(2,ROUNDUP($B$5/3,0)))')
        ws.cell(row=r, column=9, value=f'=IF(OR(G{r}="",H{r}=""),"",IF(G{r}>=H{r},"Yes","No"))')
        ws.cell(row=r, column=10, value=f'=IF(OR(B{r}<>"Board",C{r}=""),"",IFERROR(C{r}-LOOKUP(2,1/(($B$7:B{r-1}="Board")*ISNUMBER($C$7:C{r-1})),$C$7:C{r-1}),""))')
        ws.cell(row=r, column=11, value=f'=IF(J{r}="","",IF(J{r}<=120,"OK","More than 120 days"))')
        ws.cell(row=r, column=12, value=f'=IF(C{r}="","",C{r}+15)')
        ws.cell(row=r, column=13, value=f'=IF(C{r}="","",C{r}+30)')
    style_rows(ws, 8, 207, 15, formula_cols=(1, 5, 6, 8, 9, 10, 11, 12, 13), date_cols=(3, 4, 12, 13, 14), wrap_cols=(15,))
    for r in range(8, 208):
        for c in (5, 8, 10):
            ws.cell(row=r, column=c).number_format = '0'
    dv = DataValidation(type="list", formula1='"Board,Audit Committee,Nomination and Remuneration Committee,Stakeholders Relationship Committee,CSR Committee,Other committee"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("B8:B207")
    red = PatternFill("solid", bgColor="F4B6AD", fgColor="F4B6AD")
    for col in ("F", "I", "K"):
        ws.conditional_formatting.add(f"{col}8:{col}207", FormulaRule(formula=[f'OR({col}8="Shorter than 7 days",{col}8="No",{col}8="More than 120 days")'], fill=red))
    ws.freeze_panes = "A8"
    readme(wb, "Board and committee meeting tracker", "Sections 173 and 174, Companies Act, 2013; Secretarial Standard 1",
           ["1. Enter the total number of directors in cell B5. Quorum is worked out as one-third of the total or two directors, whichever is higher, with any fraction rounded up (Section 174).",
            "2. Add one row per meeting, in date order. Pick 'Board' for Board meetings so the 120-day gap check works.",
            "3. Notice must be at least seven days (Section 173(3)). A shorter notice needs the conditions in that section: the flag is only a reminder.",
            "4. Draft minutes should go to the directors within 15 days and be signed within 30 days of the meeting (Secretarial Standard 1).",
            "5. A company needs at least four Board meetings in a year, with not more than 120 days between two meetings. One person companies, small companies and dormant companies have a different rule."])
    save(wb, "board-meeting-tracker.xlsx", "Board and committee meeting tracker", "Notice, quorum, the 120-day gap, the four-meetings-a-year test and the dates for draft and signed minutes.", "Sections 173 and 174; Secretarial Standard 1", ["Meetings"])

    # 3. Directors' attendance register --------------------------------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    title(ws, "Directors' attendance register", "Type the meeting dates across the top. Mark P (present), A (absent) or L (leave granted).")
    ws["A3"], ws["B3"] = "Company name", None
    ws["A3"].font = Font(bold=True)
    ws["B3"].fill = PatternFill("solid", fgColor=INPUT)
    ws["C3"] = "Financial year"
    ws["C3"].font = Font(bold=True)
    ws["D3"].fill = PatternFill("solid", fgColor=INPUT)
    head(ws, 5, ["Sr", "Name of director", "DIN"] + [f"Meeting {i}" for i in range(1, 13)] + ["Meetings held", "Attended", "On leave", "% attended"], [5, 30, 12] + [11] * 12 + [10, 10, 10, 11])
    ws.cell(row=4, column=3, value="Meeting date  >").font = Font(italic=True, color="4B5A70")
    for c in range(4, 16):
        cell = ws.cell(row=4, column=c)
        cell.number_format = DATEF
        cell.fill = PatternFill("solid", fgColor=INPUT)
        cell.border = BOX
    for i in range(20):
        r = 6 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=16, value="=COUNT($D$4:$O$4)")
        ws.cell(row=r, column=17, value=f'=COUNTIF(D{r}:O{r},"P")')
        ws.cell(row=r, column=18, value=f'=COUNTIF(D{r}:O{r},"L")')
        ws.cell(row=r, column=19, value=f'=IF(P{r}=0,"",Q{r}/P{r})')
        ws.cell(row=r, column=19).number_format = "0%"
    style_rows(ws, 6, 25, 19, formula_cols=(1, 16, 17, 18, 19))
    for c in range(1, 20):
        pass
    for r in range(6, 26):
        ws.cell(row=r, column=19).number_format = "0%"
        for c in range(4, 16):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
    ws.cell(row=26, column=2, value="Directors present").font = Font(bold=True)
    for c in range(4, 16):
        L = get_column_letter(c)
        ws.cell(row=26, column=c, value=f'=IF({L}$4="","",COUNTIF({L}6:{L}25,"P"))').alignment = Alignment(horizontal="center")
    dv = DataValidation(type="list", formula1='"P,A,L"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("D6:O25")
    ws.conditional_formatting.add("D6:O25", FormulaRule(formula=['D6="P"'], fill=PatternFill("solid", bgColor="B7E1C4", fgColor="B7E1C4")))
    ws.conditional_formatting.add("D6:O25", FormulaRule(formula=['D6="A"'], fill=PatternFill("solid", bgColor="F4B6AD", fgColor="F4B6AD")))
    ws.conditional_formatting.add("D6:O25", FormulaRule(formula=['D6="L"'], fill=PatternFill("solid", bgColor="FFE08A", fgColor="FFE08A")))
    ws.freeze_panes = "D6"
    readme(wb, "Directors' attendance register", "Section 118 and Secretarial Standard 1 (attendance is recorded in the minutes); Regulation 17, SEBI LODR for listed companies",
           ["1. Type the date of each meeting in row 4 (up to 12 meetings).", "2. List each director's name and DIN, then mark P, A or L under each meeting.", "3. The columns on the right count the meetings held, attended and on leave, and give the percentage of attendance."])
    save(wb, "directors-attendance-register.xlsx", "Directors' attendance register", "A grid of directors against meetings, with counts and the percentage of attendance.", "Section 118; Secretarial Standard 1", ["Attendance"])

    GRP[0] = "Statutory registers"
    # generic register helper ------------------------------------------------------------
    def register(fn, name, blurb, basis, cols, readme_lines, rows=300, validations=(), extra_top=None):
        wb = Workbook()
        ws = wb.active
        ws.title = "Register"
        title(ws, name, basis)
        ws["A3"], ws["C3"] = "Company name", None
        ws["A3"].font = Font(bold=True)
        ws["B3"].fill = PatternFill("solid", fgColor=INPUT)
        ws.cell(row=3, column=3, value="CIN").font = Font(bold=True)
        ws.cell(row=3, column=4).fill = PatternFill("solid", fgColor=INPUT)
        top = 5
        if extra_top:
            for k, (lab, f) in enumerate(extra_top):
                ws.cell(row=3 + k, column=6, value=lab).font = Font(bold=True)
                ws.cell(row=3 + k, column=8, value=f)
        hdr = ["Sr"] + [c[0] for c in cols]
        head(ws, top, hdr, [5] + [c[1] for c in cols])
        first, last = top + 1, top + rows
        formula_cols, date_cols, num_cols, wrap_cols = [1], [], [], []
        for j, c in enumerate(cols, 2):
            kind = c[2]
            if kind == "date":
                date_cols.append(j)
            elif kind == "num":
                num_cols.append(j)
            elif kind == "f":
                formula_cols.append(j)
            elif kind == "wrap":
                wrap_cols.append(j)
        for i in range(rows):
            r = first + i
            ws.cell(row=r, column=1, value=i + 1)
            for j, c in enumerate(cols, 2):
                if c[2] == "f":
                    ws.cell(row=r, column=j, value=c[3].replace("{r}", str(r)).replace("{first}", str(first)).replace("{last}", str(last)))
        style_rows(ws, first, last, len(cols) + 1, formula_cols=tuple(formula_cols), date_cols=tuple(date_cols), num_cols=tuple(num_cols), wrap_cols=tuple(wrap_cols))
        for j in range(2, len(cols) + 2):
            if cols[j - 2][2] == "f":
                h = cols[j - 2][0]
                fmt = DATEF if h.startswith(("Last date", "Certificate due", "Transfer to", "Expected date")) else "0.00%" if h.startswith("%") else "#,##0.00" if "(Rs.)" in h else None
                if fmt:
                    for r in range(first, last + 1):
                        ws.cell(row=r, column=j).number_format = fmt
        for j in range(2, len(cols) + 2):
            if cols[j - 2][2] == "pct":
                for r in range(first, last + 1):
                    ws.cell(row=r, column=j).number_format = "0.00%"
        for col_idx, options in validations:
            L = get_column_letter(col_idx)
            dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"{L}{first}:{L}{last}")
        ws.freeze_panes = ws.cell(row=first, column=3)
        ws.auto_filter.ref = f"A{top}:{get_column_letter(len(cols) + 1)}{last}"
        readme(wb, name, basis, readme_lines)
        save(wb, fn, name, blurb, basis, ["Register"])
        return ws

    register("register-of-directors-and-kmp.xlsx", "Register of directors and key managerial personnel", "The particulars the Act requires for each director and each key managerial person, with a column that shows who is in office now.",
             "Section 170; Rule 17, Companies (Appointment and Qualification of Directors) Rules, 2014",
             [("Name", 28, "text"), ("Father's, mother's or spouse's name", 26, "text"), ("DIN or PAN", 16, "text"), ("Date of birth", 13, "date"), ("Nationality", 13, "text"), ("Residential address", 36, "wrap"),
              ("Occupation", 16, "text"), ("Position", 24, "text"), ("Date of appointment", 14, "date"), ("Date of cessation", 14, "date"), ("Other directorships and designated partnerships", 36, "wrap"),
              ("Shares or debentures held in the company", 22, "text"), ("In office now?", 12, "f", '=IF(OR(J{r}="",K{r}<>""),IF(J{r}="","","No"),"Yes")'), ("Remarks", 26, "wrap")],
             ["1. Add one row for each director and each key managerial person (Managing Director, Whole-time Director, CEO, CFO, Company Secretary), in the order of appointment.", "2. Fill in the date of cessation when the person leaves. 'In office now?' then changes to No.",
              "3. Keep the register at the registered office. File Form DIR-12 within 30 days of an appointment or a change."],
             validations=[(9, ["Managing Director", "Whole-time Director", "Director", "Independent Director", "Nominee Director", "Additional Director", "Alternate Director", "Chief Executive Officer", "Chief Financial Officer", "Company Secretary"])])

    register("register-of-members.xlsx", "Register of members", "The members, their shares and their certificates, with the percentage each member holds.", "Section 88; Rule 3, Companies (Management and Administration) Rules, 2014 (Form MGT-1)",
             [("Folio no.", 11, "text"), ("Name of member", 28, "text"), ("Address", 36, "wrap"), ("Email", 24, "text"), ("PAN", 13, "text"), ("Status", 18, "text"), ("Nationality", 13, "text"), ("Date of becoming a member", 14, "date"),
              ("Date of ceasing to be a member", 14, "date"), ("Class of shares", 14, "text"), ("Number of shares held", 13, "num"), ("Distinctive numbers from", 14, "text"), ("Distinctive numbers to", 14, "text"),
              ("Certificate no.", 12, "text"), ("Amount paid up per share", 14, "num"), ("Nominee", 22, "text"), ("% of total shares", 11, "f", '=IF(OR(L{r}="",$H$3=0),"",L{r}/$H$3)'), ("Remarks", 24, "wrap")],
             ["1. Add one row for each holding of each member. Fill in the date of ceasing when shares are transferred away or the member leaves.", "2. The total on the top right adds up the 'Number of shares held' column. Rows with a date of ceasing are still counted, so clear them or move them to a sheet of past members.",
              "3. Keep the register at the registered office, or at another place in the same city if the members have agreed to it by special resolution."],
             extra_top=[("Total shares in register", '=SUM(L6:L305)')])

    register("register-of-charges.xlsx", "Register of charges", "Every charge on the company's property, with the due date for filing it and the status of satisfaction.", "Section 85; Rule 10, Companies (Registration of Charges) Rules, 2014 (Form CHG-7)",
             [("Charge ID (from CHG-1)", 16, "text"), ("Date of creation", 14, "date"), ("Date of change", 14, "date"), ("Amount secured (Rs.)", 16, "num"), ("Short description of the property charged", 38, "wrap"), ("Terms of repayment", 28, "wrap"),
              ("Name and address of the charge holder", 34, "wrap"), ("Last date to file CHG-1", 14, "f", '=IF(C{r}="","",C{r}+30)'), ("CHG-1 filed on", 14, "date"), ("Date of satisfaction", 14, "date"), ("CHG-4 filed on", 14, "date"),
              ("Status", 18, "f", '=IF(C{r}="","",IF(L{r}<>"","Satisfied",IF(K{r}<>"","CHG-4 pending",IF(AND(J{r}="",TODAY()>I{r}),"CHG-1 overdue","Subsisting"))))'), ("Remarks", 24, "wrap")],
             ["1. Add a row when a charge is created or changed. The last date to file Form CHG-1 is 30 days from the date of creation.", "2. When the debt is paid, type the date of satisfaction and the date you filed Form CHG-4. The status changes to Satisfied.",
              "3. The register must be kept at the registered office and be open to inspection."],
             rows=200)

    register("register-of-loans-guarantees-investments.xlsx", "Register of loans, guarantees, securities and investments", "Every loan, guarantee, security and investment the company makes, with the approvals and a running total.", "Section 186(9); Rule 12, Companies (Meetings of Board and its Powers) Rules, 2014 (Form MBP-2)",
             [("Date of making", 14, "date"), ("Type", 18, "text"), ("Name of the body corporate or person", 30, "text"), ("Amount (Rs.)", 16, "num"), ("Purpose", 28, "wrap"), ("Date of Board resolution", 14, "date"),
              ("Date of special resolution (if needed)", 16, "date"), ("Rate of interest (% a year)", 12, "num"), ("Date of repayment or release", 14, "date"), ("Total so far (Rs.)", 16, "f", '=IF(E{r}="","",SUM($E${first}:E{r}))'), ("Remarks", 26, "wrap")],
             ["1. Add a row for each transaction in the order it was made. Pick the type from the list.", "2. Type the dates of the Board resolution and, where needed, the special resolution of the members.", "3. The 'Total so far' column helps you test the limits in Section 186(2). Compare it with 60% of the paid-up capital, free reserves and securities premium, or 100% of free reserves and securities premium, whichever is higher."],
             validations=[(3, ["Loan", "Guarantee", "Security", "Investment"])])

    register("register-of-related-party-contracts.xlsx", "Register of contracts with related parties and directors", "Contracts and arrangements in which a director is interested, with the approval dates.", "Section 189; Rule 16, Companies (Meetings of Board and its Powers) Rules, 2014 (Form MBP-4)",
             [("Date of entry", 13, "date"), ("Name of the related party", 28, "text"), ("Relationship", 22, "text"), ("Nature of the contract or arrangement", 32, "wrap"), ("Date of the contract", 13, "date"), ("Period", 14, "text"),
              ("Main terms", 32, "wrap"), ("Value (Rs.)", 16, "num"), ("Name of the interested director or key managerial person", 28, "text"), ("Date of Board approval", 14, "date"), ("Date of the members' resolution (if needed)", 16, "date"),
              ("Date the director disclosed his interest (MBP-1)", 16, "date"), ("Remarks", 24, "wrap")],
             ["1. Enter each contract or arrangement under Section 188 or in which a director is interested.", "2. Record the date of the Board approval and, where the limits in Rule 15 are crossed, the members' resolution.", "3. A related party who votes cannot be counted for that resolution."])

    GRP[0] = "Filings, meetings and planners"
    # 9. AGM planner ---------------------------------------------------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "AGM planner"
    title(ws, "AGM planner", "Enter the year end and the AGM date. The key dates are worked out.")
    inputs(ws, [("Company name", None), ("Financial year ends on", DATEF), ("Date of the AGM", DATEF)], 1, 2, 3)
    head(ws, 7, ["Step", "What to do", "Date", "How worked out"], [5, 54, 16, 54])
    plan = [
        ("Last date for the AGM", '=IF($B$4="","",EDATE($B$4,6))', "6 months from the year end (Section 96)"),
        ("Latest date to send the notice", '=IF($B$5="","",$B$5-22)', "21 clear days: the day of the notice and the day of the meeting are not counted"),
        ("Latest date to send the Annual Report", '=IF($B$5="","",$B$5-22)', "Goes with the notice (Section 136)"),
        ("Latest date for a proxy form to reach the company", '=IF($B$5="","",$B$5-2)', "At least 48 hours before the meeting (Section 105)"),
        ("Draft minutes to the directors", '=IF($B$5="","",$B$5+15)', "Within 15 days (Secretarial Standard 2)"),
        ("Minutes signed by the Chairman", '=IF($B$5="","",$B$5+30)', "Within 30 days (Secretarial Standard 2)"),
        ("File ADT-1", '=IF($B$5="","",$B$5+15)', "15 days from the AGM"),
        ("File AOC-4", '=IF($B$5="","",$B$5+30)', "30 days from the AGM"),
        ("File MGT-7", '=IF($B$5="","",$B$5+60)', "60 days from the AGM"),
    ]
    for i, (a, f, how) in enumerate(plan):
        r = 8 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=a)
        ws.cell(row=r, column=3, value=f)
        ws.cell(row=r, column=4, value=how)
    style_rows(ws, 8, 8 + len(plan) - 1, 4, formula_cols=(3,), date_cols=(3,), wrap_cols=(2, 4))
    for r in range(8, 8 + len(plan)):
        for c in (1, 2, 4):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFFFFF")
    ws.freeze_panes = "A8"
    readme(wb, "AGM planner", "Sections 96, 101, 105, 136 and 137; Secretarial Standards 2",
           ["1. Fill in the financial year end and the date you plan to hold the AGM.", "2. Work back from the AGM date for the notice and forward from it for the filings.", "3. A shorter notice is possible with the consent of members holding at least 95% of the voting power. This planner uses the normal 21 clear days."])
    save(wb, "agm-planner.xlsx", "AGM planner", "Work back from the AGM date for the notice and the Annual Report, and forward for the minutes and the filings.", "Sections 96, 101, 105, 136, 137", ["AGM planner"])


    # ---- more registers ---------------------------------------------------------------
    GRP[0] = "Statutory registers"
    register("register-of-allotments.xlsx", "Register of allotments", "Each allotment of shares, with the consideration worked out and the last date to file Form PAS-3.", "Sections 39 and 42; Rules 12 and 14 of the Share Capital and Debentures Rules and the Prospectus and Allotment Rules",
             [("Date of allotment", 14, "date"), ("Type of issue", 20, "text"), ("Name of allottee", 28, "text"), ("Address", 32, "wrap"), ("PAN", 13, "text"), ("Class of shares", 14, "text"), ("Number of shares", 12, "num"), ("Face value (Rs.)", 12, "num"),
              ("Issue price per share (Rs.)", 14, "num"), ("Total consideration (Rs.)", 16, "f", '=IF(OR(H{r}="",J{r}=""),"",H{r}*J{r})'), ("Mode of payment", 16, "text"), ("Distinctive numbers from", 14, "text"), ("Distinctive numbers to", 14, "text"), ("Certificate no.", 12, "text"),
              ("Date of Board resolution", 14, "date"), ("Date of members' approval", 14, "date"), ("Last date to file PAS-3", 14, "f", '=IF(B{r}="","",B{r}+IF(C{r}="Private placement",15,30))'),
              ("PAS-3 filed on", 14, "date"), ("Remarks", 24, "wrap")],
             ["1. Add a row for each allottee in each allotment. Pick the type of issue from the list.", "2. The last date to file Form PAS-3 is 15 days from allotment for a private placement and 30 days for other issues. Check the current Rules.",
              "3. Issue share certificates within two months of allotment (Section 56(4)) and enter each holder in the Register of Members."],
             validations=[(3, ["Rights issue", "Bonus issue", "Private placement", "Preferential allotment", "Employee stock option", "Sweat equity", "Subscribers to the Memorandum", "Other"])])
    register("share-transfer-register.xlsx", "Share transfer register", "Transfers and transmissions of shares, with the date by which the new certificate is due.", "Sections 56 and 58; Secretarial Standard on transfer; Rule 7 (transfer and transmission)",
             [("Transfer no.", 11, "text"), ("Date SH-4 was received", 14, "date"), ("Date of Board approval", 14, "date"), ("Name of transferor", 26, "text"), ("Name of transferee", 26, "text"), ("Folio of transferor", 11, "text"), ("Folio of transferee", 11, "text"),
              ("Class of shares", 13, "text"), ("Number of shares", 12, "num"), ("Distinctive numbers from", 13, "text"), ("Distinctive numbers to", 13, "text"), ("Old certificate no.", 12, "text"), ("New certificate no.", 12, "text"),
              ("Price per share (Rs.)", 13, "num"), ("Total value (Rs.)", 15, "f", '=IF(OR(J{r}="",O{r}=""),"",J{r}*O{r})'), ("Stamp duty paid (Rs.)", 13, "num"), ("Certificate due by (one month)", 14, "f", '=IF(C{r}="","",C{r}+30)'),
              ("New certificate issued on", 14, "date"), ("Status", 16, "f", '=IF(C{r}="","",IF(S{r}<>"","Issued",IF(TODAY()>R{r},"Certificate overdue","Pending")))'), ("Remarks", 24, "wrap")],
             ["1. Add a row for each transfer, and for each transmission (mention it in the remarks).", "2. The company must deliver the new certificate within one month of receiving the instrument of transfer (Section 56(4)(c)). The due date is worked out for you.",
              "3. Check that the SH-4 is properly stamped. Stamp duty on transfers differs by State and by the type of company."])
    register("register-of-debenture-holders.xlsx", "Register of debenture holders and other security holders", "Holders of debentures and other securities, with the amount and the date of redemption.", "Section 88(1)(b) and (c); Rule 3 (Form MGT-2)",
             [("Folio no.", 11, "text"), ("Name of holder", 28, "text"), ("Address", 32, "wrap"), ("Email", 24, "text"), ("PAN", 13, "text"), ("Date of becoming a holder", 14, "date"), ("Series or class", 16, "text"), ("Number of debentures", 12, "num"),
              ("Face value each (Rs.)", 13, "num"), ("Amount (Rs.)", 15, "f", '=IF(OR(I{r}="",J{r}=""),"",I{r}*J{r})'), ("Distinctive numbers", 18, "text"), ("Certificate no.", 12, "text"), ("Rate of interest (%)", 11, "num"),
              ("Date of redemption", 14, "date"), ("Date of ceasing to be a holder", 14, "date"), ("Remarks", 24, "wrap")],
             ["1. Add a row for each holding of each debenture holder.", "2. The amount is worked out from the number and the face value.", "3. The register is kept at the registered office and open to inspection."],
             extra_top=[("Total amount of debentures (Rs.)", '=SUM(K6:K305)')])
    register("register-of-renewed-duplicate-certificates.xlsx", "Register of renewed and duplicate share certificates", "Duplicate and renewed share certificates, with the reason and the Board approval.", "Section 46; Rule 6 (Form SH-2)",
             [("Date of issue", 13, "date"), ("New certificate no.", 13, "text"), ("Name of holder", 28, "text"), ("Folio no.", 11, "text"), ("Number of shares", 12, "num"), ("Distinctive numbers", 18, "text"), ("Original certificate no.", 13, "text"),
              ("Reason", 22, "text"), ("Date of Board approval", 14, "date"), ("Documents received (FIR, affidavit, indemnity)", 28, "wrap"), ("Original cancelled?", 12, "text"), ("Date entered in the Register of Members", 16, "date"), ("Signed by (director or secretary)", 22, "text"), ("Remarks", 24, "wrap")],
             ["1. Add a row each time a duplicate or renewed certificate is issued.", "2. Mark the new certificate 'Duplicate' and record the Board resolution that approved it.", "3. Keep the register at the registered office."],
             validations=[(9, ["Lost", "Destroyed", "Torn or defaced", "Pages for transfer used up"]), (12, ["Yes", "No", "Not available"])])
    register("register-of-dividends.xlsx", "Register of dividends", "Dividends declared and paid, with the dates for the unpaid dividend account and for transfer to the IEPF.", "Sections 123, 124 and 125, Companies Act, 2013; Investor Education and Protection Fund Rules",
             [("Financial year", 12, "text"), ("Type", 12, "text"), ("Date of declaration", 14, "date"), ("Rate per share (Rs.)", 12, "num"), ("Total declared (Rs.)", 16, "num"), ("Last date to pay (30 days)", 14, "f", '=IF(D{r}="","",D{r}+30)'),
              ("Amount paid to members (Rs.)", 16, "num"), ("Tax deducted and deposited (Rs.)", 16, "num"), ("Unpaid amount (Rs.)", 16, "f", '=IF(F{r}="","",F{r}-H{r}-I{r})'), ("Transfer to Unpaid Dividend Account due by (37 days)", 18, "f", '=IF(D{r}="","",D{r}+37)'),
              ("Date transferred to the Unpaid Dividend Account", 16, "date"), ("Transfer to IEPF due after 7 years", 16, "f", '=IF(L{r}="","",EDATE(L{r},84))'), ("Remarks", 24, "wrap")],
             ["1. Add a row for each dividend. Pick Interim or Final.", "2. An interim dividend must be moved to a separate bank account within five days of declaration (Section 123(4)). Final dividend and interim dividend must be paid within thirty days.",
              "3. Any dividend left unpaid after thirty days goes to the Unpaid Dividend Account within seven days (Section 124). Amounts unclaimed for seven years go to the Investor Education and Protection Fund (Section 125)."],
             validations=[(3, ["Interim", "Final", "Special"])])
    GRP[0] = "Practice"
    register("register-of-articled-assistants.xlsx", "Register of articled assistants", "Articleship records, with the ICAI forms and the dates they fall due.", "Chartered Accountants Regulations, 1988 (ICAI Forms 102, 103, 108 and 109)",
             [("Name of the articled assistant", 28, "text"), ("Registration no.", 15, "text"), ("Date of registration", 14, "date"), ("Date of commencement", 14, "date"), ("Period in months", 10, "num"), ("Expected date of completion", 15, "f", '=IF(E{r}="","",EDATE(E{r},F{r}))'),
              ("Stipend per month (Rs.)", 14, "num"), ("Form 102 filed on", 14, "date"), ("Form 103 filed on", 14, "date"), ("Form 108 filed on", 14, "date"), ("Form 109 filed on", 14, "date"), ("Leave taken (days)", 11, "num"), ("Remarks", 24, "wrap")],
             ["1. Add a row for each articled assistant. Enter the period of articleship in months.", "2. The expected date of completion is the commencement date plus the period. Add the leave beyond the allowed days yourself.", "3. Check the current ICAI regulations for the forms and the time limits."])

    # ---- attendance registers -----------------------------------------------------------
    GRP[0] = "Filings, meetings and planners"
    wb = Workbook()
    wb.remove(wb.active)
    for nm, ttl, cols, w in [("Board meeting", "Board meeting attendance register", ["Name of director", "DIN", "Present in person or by video", "Signature"], [30, 14, 22, 26]),
                             ("AGM directors", "General meeting: attendance of directors", ["Name of director", "DIN", "Position", "Signature"], [30, 14, 22, 26]),
                             ("AGM members", "General meeting: attendance of members", ["Name of member or proxy", "Folio or DP and client ID", "Number of shares", "In person or by proxy", "Signature"], [30, 20, 14, 18, 26])]:
        ws = wb.create_sheet(nm)
        title(ws, ttl, "Print this sheet and have each person sign in at the meeting.")
        ws["A3"], ws["A4"], ws["A5"] = "Company", "Date and time of the meeting", "Place"
        for r in (3, 4, 5):
            ws.cell(row=r, column=1).font = Font(bold=True)
            ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=INPUT)
        head(ws, 7, ["Sr"] + cols, [5] + w)
        n = 40 if nm == "AGM members" else 15
        for i in range(n):
            ws.cell(row=8 + i, column=1, value=i + 1)
        style_rows(ws, 8, 7 + n, len(cols) + 1, formula_cols=(1,))
        for r in range(8, 8 + n):
            ws.row_dimensions[r].height = 26
        ws.cell(row=9 + n, column=2, value="Total present").font = Font(bold=True)
        ws.cell(row=9 + n, column=3, value=f'=COUNTA(B8:B{7 + n})')
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToHeight = 0
    readme(wb, "Meeting attendance registers", "Secretarial Standards 1 and 2",
           ["1. Pick the sheet for the meeting. Fill in the company, the date and the place.", "2. Print it, and have each person sign against his or her name.", "3. The 'Total present' line counts the names written in. Use it to check the quorum before the meeting begins."])
    save(wb, "meeting-attendance-registers.xlsx", "Meeting attendance registers", "Print-ready sign-in sheets for Board meetings, and for directors and members at general meetings.", "Secretarial Standards 1 and 2", ["Board meeting", "AGM directors", "AGM members"])

    # ---- checklists -----------------------------------------------------------------------
    GRP[0] = "Checklists"

    def checklist(fn, name, blurb, basis, inputs_list, items, lines):
        wb = Workbook()
        ws = wb.active
        ws.title = "Checklist"
        title(ws, name, "Fill in the dates in the yellow cells. The due dates and the status work themselves out.")
        ws["A3"] = "Company name"
        ws["A3"].font = Font(bold=True)
        ws["B3"].fill = PatternFill("solid", fgColor=INPUT)
        for i, lab in enumerate(inputs_list):
            r = 4 + i
            ws.cell(row=r, column=1, value=lab).font = Font(bold=True)
            c = ws.cell(row=r, column=2)
            c.number_format = DATEF
            c.fill = PatternFill("solid", fgColor=INPUT)
            c.border = BOX
        top = 5 + len(inputs_list) + 1
        head(ws, top, ["Sr", "Step", "Law", "Time limit", "Due date", "Who", "Done?", "Date done", "Status", "Remarks"], [5, 56, 24, 30, 14, 16, 14, 14, 18, 28])
        for i, (stp, law, lim, f, who) in enumerate(items):
            r = top + 1 + i
            ws.cell(row=r, column=1, value=i + 1)
            ws.cell(row=r, column=2, value=stp)
            ws.cell(row=r, column=3, value=law)
            ws.cell(row=r, column=4, value=lim)
            if f:
                ws.cell(row=r, column=5, value=f)
            ws.cell(row=r, column=6, value=who)
            ws.cell(row=r, column=9, value=f'=IF(G{r}="Yes","Done",IF(G{r}="Not applicable","Not applicable",IF(E{r}="","To do",IF(E{r}<TODAY(),"Overdue",IF(E{r}-TODAY()<=15,"Due in "&(E{r}-TODAY())&" days","Upcoming")))))')
        last = top + len(items) + 3
        style_rows(ws, top + 1, last, 10, formula_cols=(1, 5, 9), date_cols=(5, 8), wrap_cols=(2, 3, 4, 10))
        for r in range(top + 1, top + 1 + len(items)):
            for c in (2, 3, 4, 6):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFFFFF")
        dv = DataValidation(type="list", formula1='"Yes,No,Not applicable"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"G{top + 1}:G{last}")
        ws.conditional_formatting.add(f"I{top + 1}:I{last}", FormulaRule(formula=[f'I{top + 1}="Overdue"'], fill=PatternFill("solid", bgColor="F4B6AD", fgColor="F4B6AD"), font=Font(bold=True, color="8A1C0E")))
        ws.conditional_formatting.add(f"I{top + 1}:I{last}", FormulaRule(formula=[f'LEFT(I{top + 1},3)="Due"'], fill=PatternFill("solid", bgColor="FFE08A", fgColor="FFE08A")))
        ws.conditional_formatting.add(f"I{top + 1}:I{last}", FormulaRule(formula=[f'I{top + 1}="Done"'], fill=PatternFill("solid", bgColor="B7E1C4", fgColor="B7E1C4")))
        ws.freeze_panes = ws.cell(row=top + 1, column=3)
        readme(wb, name, basis, lines)
        save(wb, fn, name, blurb, basis, ["Checklist"])

    checklist("incorporation-first-90-days-checklist.xlsx", "New company: first 90 days checklist", "The company-law steps after incorporation, with due dates from the date of incorporation.", "Sections 10A, 12, 139(6), 173(2), 56(4), 184 and 88, Companies Act, 2013",
              ["Date of incorporation"],
              [("Have a registered office that can receive communications", "Section 12(1)", "From the 15th day of incorporation", '=IF($B$4="","",$B$4+15)', "Director"),
               ("File the notice of the registered office (INC-22)", "Section 12(2)", "Within 30 days of incorporation", '=IF($B$4="","",$B$4+30)', "Company Secretary"),
               ("Hold the first Board meeting", "Section 173(2)", "Within 30 days of incorporation", '=IF($B$4="","",$B$4+30)', "Chairman"),
               ("Directors give notice of interest (MBP-1) at the first meeting", "Section 184", "At the first Board meeting", '=IF($B$4="","",$B$4+30)', "Directors"),
               ("Appoint the first auditor by the Board", "Section 139(6)", "Within 30 days of incorporation", '=IF($B$4="","",$B$4+30)', "Board"),
               ("If the Board does not, appoint the first auditor at a general meeting", "Section 139(6)", "Within 90 days of incorporation", '=IF($B$4="","",$B$4+90)', "Members"),
               ("File Form ADT-1 for the first auditor", "Section 139; Rule 4", "Within 15 days of the appointment", None, "Company Secretary"),
               ("Open the company's bank account and receive the subscribers' money", "Board resolution", "Before filing the INC-20A", None, "Director"),
               ("Open the Register of Members and the Register of Directors", "Sections 88 and 170", "From incorporation", None, "Company Secretary"),
               ("Issue share certificates to the subscribers (Form SH-1)", "Section 56(4)(a)", "Within 2 months of incorporation", '=IF($B$4="","",EDATE($B$4,2))', "Company Secretary"),
               ("File the declaration of commencement of business (INC-20A)", "Section 10A", "Within 180 days of incorporation", '=IF($B$4="","",$B$4+180)', "Company Secretary"),
               ("Fix the financial year and tell the tax authorities", "Section 2(41)", "At the first Board meeting", None, "Board"),
               ("Register for GST, Udyam and other licences if they apply", "Other laws", "As needed", None, "Director"),
               ("Put the statutory registers, the minute books and the seal (if used) in place", "Sections 118 and 88", "From incorporation", None, "Company Secretary")],
              ["1. Type the date of incorporation in cell B4.", "2. Each step shows its due date. Pick Yes when it is done and type the date.", "3. Dates come from the Act. Check whether any extension or exemption applies to your company."])

    checklist("private-placement-funding-checklist.xlsx", "Funding round by private placement: checklist", "The steps and dates for raising money by issuing shares to selected investors.", "Sections 42, 62 and 117, Companies Act, 2013; Rules 13 and 14, Share Capital and Debentures Rules and Prospectus and Allotment Rules",
              ["Date of Board meeting approving the issue", "Date of the EGM", "Date money was received", "Date of allotment"],
              [("Get a valuation of the shares from a registered valuer", "Rule 13", "Before the Board and members approve the price", None, "Company"),
               ("Hold the Board meeting to approve the issue and call the EGM", "Sections 42, 62", "On the date entered above", '=IF($B$4="","",$B$4)', "Board"),
               ("Send the EGM notice with the explanatory statement", "Sections 101 and 102", "At least 21 clear days before the EGM", '=IF($B$5="","",$B$5-22)', "Company Secretary"),
               ("Pass the special resolution for the preferential offer", "Sections 42(2) and 62(1)(c)", "On the date of the EGM", '=IF($B$5="","",$B$5)', "Members"),
               ("File the special resolution (MGT-14)", "Section 117", "Within 30 days of the EGM", '=IF($B$5="","",$B$5+30)', "Company Secretary"),
               ("Make the offers in Form PAS-4 to named persons, in writing", "Rule 14", "Within the validity of the special resolution", None, "Company"),
               ("Keep the record of offers in Form PAS-5", "Rule 14", "At the time of the offer", None, "Company Secretary"),
               ("Check that the number of persons offered is within the limit for the year", "Section 42(2); Rule 14", "200 persons in a year, excluding qualified institutional buyers and employees under a stock option scheme", None, "Company Secretary"),
               ("Receive the money through a bank, from the investor's own account", "Rule 14", "Before the allotment", None, "Investor"),
               ("Allot the shares", "Section 42(6)", "Within 60 days of receiving the money", '=IF($B$6="","",$B$6+60)', "Board"),
               ("If the shares are not allotted, return the money", "Section 42(6)", "Within 15 days after the 60 days end", '=IF($B$6="","",$B$6+75)', "Company"),
               ("File the return of allotment (PAS-3)", "Rule 14", "Within 15 days of the allotment", '=IF($B$7="","",$B$7+15)', "Company Secretary"),
               ("Issue the share certificates (Form SH-1)", "Section 56(4)", "Within 2 months of the allotment", '=IF($B$7="","",EDATE($B$7,2))', "Company Secretary"),
               ("Update the Register of Members and the cap table", "Section 88", "On allotment", None, "Company Secretary"),
               ("If the investor is outside India, file the foreign investment return (FC-GPR)", "FEMA rules", "Within 30 days of the allotment", '=IF($B$7="","",$B$7+30)', "Company Secretary")],
              ["1. Fill in the four dates at the top as the round goes on.", "2. Each step shows its due date. Pick Yes when it is done and type the date.", "3. The valuation, the offer and the deadlines depend on the type of company and the kind of investor. Check each rule before you rely on it."])

    order = ["Statutory registers", "Filings, meetings and planners", "Checklists", "Practice"]
    sections = ""
    for gname in order:
        items = [m for m in made if m["group"] == gname]
        if items:
            sections += f"<h2>{gname}</h2><div class=\"xl\">" + "".join(
                f'<div class="card"><b>{m["name"]}</b><span>{m["blurb"]}</span><i>Based on: {m["basis"]}</i><a class="btn" href="{m["file"]}" download>Download .xlsx</a></div>' for m in items) + "</div>"
    body = f"""<h1>Excel utilities</h1><p class="lede">Free working files for the Company Secretary's desk: statutory registers, trackers, planners and checklists. Each file opens in Excel, LibreOffice and Google Sheets. Cells you fill in are yellow; cells that work themselves out are grey-blue.</p>
{sections}
<h2>How to use them</h2><ul><li>Download a file and save it with your company's name.</li><li>Fill in the yellow cells. The dates, totals and checks update on their own.</li><li>Check each rule and date against the current law. These files do not include extensions or exemptions announced by the Ministry.</li></ul>"""
    (out / "index.html").write_text(shell("Excel utilities | Company Law Knowledge Lab", "Statutory registers, compliance trackers and planners as free Excel files.", "xl", body), encoding="utf-8")
    print(f"{len(made)} Excel files written")


if __name__ == "__main__":
    for step in (build_generators, build_excel):
        try:
            step()
        except Exception as e:
            print(f"{step.__name__} ended with an error: {e}")
    sys.exit(0)
