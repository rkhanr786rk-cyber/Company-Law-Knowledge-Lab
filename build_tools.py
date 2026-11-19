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
        out.push(RESF(`the share certificate(s) in Form SH-1 be issued within two months of the date of allotment, the Register of Members be updated, and the return of allotment in Form PAS-3 be filed with the Registrar of Companies within thirty days of the allotment.`));
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
  const g = GENS.find(x => x.id === id); if (!g) return listGens(); S_.cur = g;
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
  ${groups.map(gr => `<h2>${esc(gr)}</h2><div class="cards">${GENS.filter(g => g.group === gr).map(g => `<a class="card" href="#${g.id}"><b>${esc(g.title)}</b><span>${esc(g.short)}</span><i>${esc(g.basis)}</i></a>`).join("")}</div>`).join("")}`;
}
function route() { const id = location.hash.replace(/^#/, ""); if (id) openGen(id); else listGens(); }
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
        made.append({"file": fn, "name": name, "blurb": blurb, "basis": basis, "sheets": sheets})

    def status_rules(ws, rng, col, first):
        ref = f"{col}{first}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{ref}="Overdue"'], fill=PatternFill("solid", bgColor="F4B6AD", fgColor="F4B6AD"), font=Font(bold=True, color="8A1C0E")))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({ref},3)="Due"'], fill=PatternFill("solid", bgColor="FFE08A", fgColor="FFE08A")))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{ref}="Filed"'], fill=PatternFill("solid", bgColor="B7E1C4", fgColor="B7E1C4")))

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
          ("Allotment of shares", "PAS-3", "Section 39(4)", 30), ("Creation or change of a charge", "CHG-1 / CHG-9", "Section 77", 30),
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

    cards = "".join(
        f'<div class="card"><b>{m["name"]}</b><span>{m["blurb"]}</span><i>Based on: {m["basis"]}</i><a class="btn" href="{m["file"]}" download>Download .xlsx</a></div>' for m in made)
    body = f"""<h1>Excel utilities</h1><p class="lede">Free working files for the Company Secretary's desk: statutory registers, trackers and planners. Each file opens in Excel, LibreOffice and Google Sheets. Cells you fill in are yellow; cells that work themselves out are grey-blue.</p>
<div class="xl">{cards}</div>
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
