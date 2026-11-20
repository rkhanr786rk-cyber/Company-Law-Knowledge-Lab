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

.tbl{border-collapse:collapse;width:100%;background:var(--page);font-size:.95rem}
.tbl th,.tbl td{border:1px solid var(--rule);padding:.5rem .7rem;text-align:left;vertical-align:top}
.tbl th{background:var(--bar);color:var(--barink);font:600 .9rem var(--sans)}
.dt{border-collapse:collapse;width:100%;margin:.4rem 0 .9rem;font-size:.92rem}.dt th,.dt td{border:1px solid #444;padding:.3rem .5rem;text-align:left;vertical-align:top}.dt th{font-weight:700}
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
const TB = (rows, w) => ({ k: "TB", rows, w }); /* table: rows of cells; first row is the header */
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
/* ---------- batch 3: committees, EGM minutes, agenda, postal ballot, prescribed formats, policies ---------- */
const SG = [{ id: "sname", label: "Name of the signing director or officer" }, { id: "sdes", label: "Designation" }];
const SGN = v => signoff(v, va(v, "sname", "name"), va(v, "sdes", "designation"));
const CMS = {
  audit: { n: "Audit Committee", b: "Section 177, Companies Act, 2013; Rule 6, Companies (Meetings of Board and its Powers) Rules, 2014; Regulation 18, SEBI (LODR) Regulations, 2015 (listed companies)",
    q: "The quorum was present: two members or one-third of the members of the Committee, whichever is greater, including at least two independent directors (Regulation 18(2)(b) for a listed company).",
    items: ["Confirmation of the minutes of the last meeting", "Review of the financial statements and the auditor's report", "Review of related party transactions and the grant of omnibus approval", "Review of internal financial controls and the internal audit report", "Review of the independence, performance and remuneration of the auditors", "Any other business with the permission of the Chair"] },
  nrc: { n: "Nomination and Remuneration Committee", b: "Section 178, Companies Act, 2013; Rule 6, Companies (Meetings of Board and its Powers) Rules, 2014; Regulation 19, SEBI (LODR) Regulations, 2015 (listed companies)",
    q: "The quorum was present: two members or one-third of the members, whichever is greater, including at least one independent director (Regulation 19(2A) for a listed company).",
    items: ["Confirmation of the minutes of the last meeting", "Review of the composition of the Board and the need for new appointments", "Recommendation on the appointment, reappointment or remuneration of directors and senior management", "Evaluation of the Board, its committees and the individual directors", "Review of the nomination and remuneration policy", "Any other business with the permission of the Chair"] },
  src: { n: "Stakeholders Relationship Committee", b: "Section 178(5), Companies Act, 2013; Regulation 20, SEBI (LODR) Regulations, 2015 (listed companies)",
    q: "The quorum required for the meeting was present.",
    items: ["Confirmation of the minutes of the last meeting", "Review of the investor complaints received and resolved", "Review of share transfers, transmissions, dematerialisation and duplicate certificates", "Review of unclaimed dividends and shares due for transfer to the Investor Education and Protection Fund", "Any other business with the permission of the Chair"] },
  csr: { n: "Corporate Social Responsibility Committee", b: "Section 135, Companies Act, 2013; Companies (Corporate Social Responsibility Policy) Rules, 2014",
    q: "The quorum required for the meeting was present.",
    items: ["Confirmation of the minutes of the last meeting", "Review of the CSR policy and the annual action plan", "Review of the CSR projects, the amount spent and any unspent amount", "Recommendation of the CSR expenditure to the Board", "Any other business with the permission of the Chair"] },
  risk: { n: "Risk Management Committee", b: "Regulation 21, SEBI (LODR) Regulations, 2015 (top 1,000 listed companies by market capitalisation); Section 134(3)(n), Companies Act, 2013",
    q: "The quorum required for the meeting was present: two members or one-third of the members, whichever is higher, including at least one member of the Board (Regulation 21(3A)).",
    items: ["Confirmation of the minutes of the last meeting", "Review of the risk register and the risk management framework", "Review of cyber security and business continuity risks", "Any other business with the permission of the Chair"] },
};
const CMOPTS = Object.keys(CMS).map(k => [k, CMS[k].n]);
const cmName = v => (CMS[v.cmtype] || CMS.audit).n;

GENS.push(
  {
    id: "cmnotice", group: "Committees", title: "Notice of committee meeting", short: "Notice and agenda for an Audit, Nomination and Remuneration, Stakeholders, CSR or Risk Committee meeting.",
    basis: "Secretarial Standard 1 (applies to committee meetings); the section for each committee",
    fields: [{ id: "cmtype", label: "Committee", type: "select", opts: CMOPTS }, { id: "ndate", label: "Date of notice", type: "date" }, { id: "mdate", label: "Date of meeting", type: "date" }, { id: "mtime", label: "Time", ph: "11:00 a.m." }, { id: "place", label: "Place", ph: "Registered office" },
      { id: "agenda", label: "Agenda (one item per line; leave blank for the usual items)", type: "textarea" }].concat(SG),
    notes: ["Send the notice and the agenda with notes at least seven days before the meeting, unless the committee agrees to shorter notice as Secretarial Standard 1 allows.", "Each committee has its own rules on members and quorum. Check them for your company."],
    build(v) {
      const c = CMS[v.cmtype] || CMS.audit, ag = LN("agenda")(v), out = header(v);
      out.push(T(`NOTICE OF A MEETING OF THE ${c.n.toUpperCase()}`), R("Date: " + vd(v, "ndate", "date of notice")), L(`To all the members of the ${c.n}`), S());
      out.push(P(`**NOTICE** is hereby given that a meeting of the **${c.n}** of the Board of Directors of **${CO(v)}** will be held on **${vd(v, "mdate", "date")}** at **${va(v, "mtime", "time")}** at ${PLACE(v)}, to transact the following business:`));
      out.push(P("**AGENDA**")); NUMLIST(out, ag.length ? ag : c.items, "[item]");
      out.push(S(), P("The agenda, with notes on each item, is attached. You are requested to attend. If you cannot attend, please inform the Company Secretary in advance."), ...SGN(v));
      return out;
    },
  },
  {
    id: "cmminutes", group: "Committees", title: "Minutes of committee meeting", short: "Minutes with attendance, quorum, items and decisions, for any of five committees.",
    basis: "Secretarial Standard 1; the section for each committee",
    fields: [{ id: "cmtype", label: "Committee", type: "select", opts: CMOPTS }, { id: "mdate", label: "Date of meeting", type: "date" }, { id: "mtime", label: "Time", ph: "11:00 a.m." }, { id: "place", label: "Place", ph: "Registered office" },
      { id: "chair", label: "Chairperson of the meeting" }, { id: "present", label: "Members present (one per line)", type: "textarea" }, { id: "invitees", label: "Invitees present (one per line, optional)", type: "textarea" },
      { id: "items", label: "Items and decisions (one per line: item | decision). Blank gives the usual items.", type: "textarea", ph: "Review of related party transactions | The Committee approved the transactions placed before it" }, { id: "cend", label: "Time the meeting ended", ph: "12:30 p.m." }, { id: "sdate", label: "Date of signing the minutes", type: "date" }],
    notes: ["Minutes are signed by the Chairperson of the same meeting or the next one, within thirty days (Secretarial Standard 1).", "Replace each [decision] with what was actually decided."],
    build(v) {
      const c = CMS[v.cmtype] || CMS.audit, out = header(v);
      out.push(T(`MINUTES OF THE MEETING OF THE ${c.n.toUpperCase()} HELD ON ${vd(v, "mdate", "date").toUpperCase()}${v.mtime ? " AT " + v.mtime.toUpperCase() : ""} AT ${PLACE(v).toUpperCase()}`));
      out.push(P("**Members present**")); NUMLIST(out, LN("present")(v), "[names of members]");
      if (LN("invitees")(v).length) { out.push(P("**Invitees present**")); NUMLIST(out, LN("invitees")(v), ""); }
      out.push(S(), P(`${va(v, "chair", "name")} took the chair. The Chairperson noted that the required notice and agenda had been sent to all the members. ${c.q} The Chairperson called the meeting to order.`));
      const rows = LN("items")(v).map(x => x.split("|").map(y => y.trim())); const list = rows.length ? rows : c.items.map(x => [x, "[decision]"]);
      list.forEach((r, i) => { out.push(P(`**${i + 1}. ${r[0]}**`)); out.push(P(r[1] || "[decision]")); });
      out.push(P(`There being no other business, the meeting ended${v.cend ? " at " + v.cend : ""} with a vote of thanks to the Chair.`), S(), L("Date: " + vd(v, "sdate", "date")), L("Place: " + PLACE(v)), R("**" + va(v, "chair", "Chairperson") + "**"), R("Chairperson"));
      return out;
    },
  },
  {
    id: "agenda", group: "Meetings", title: "Agenda for a Board meeting", short: "A numbered agenda with the usual standing items and your own items.",
    basis: "Section 173; Secretarial Standard 1 (agenda and notes on agenda)",
    fields: [{ id: "bmno", label: "Meeting number or serial", ph: "1/2026-27" }, { id: "mdate", label: "Date of meeting", type: "date" }, { id: "mtime", label: "Time", ph: "11:00 a.m." }, { id: "place", label: "Place", ph: "Registered office" },
      { id: "items", label: "Your agenda items (one per line)", type: "textarea" }].concat(SG),
    notes: ["Send the agenda, with notes on each item, at least seven days before the meeting. Items of unpublished price sensitive information may be taken with shorter notice if an independent director attends (Secretarial Standard 1).", "Items that Section 179(3) says only the Board can decide must be taken at a Board meeting, not by circulation."],
    build(v) {
      const out = header(v), own = LN("items")(v);
      out.push(T(`AGENDA FOR THE MEETING OF THE BOARD OF DIRECTORS NO. ${va(v, "bmno", "number")}`), C(`To be held on ${vd(v, "mdate", "date")} at ${va(v, "mtime", "time")} at ${PLACE(v)}`), S());
      const std = ["Leave of absence, if any", "Confirmation of the minutes of the previous Board meeting", "Review of the action taken on the previous decisions", "Review of the financial results and the working of the Company", "Disclosures of interest by directors (Section 184)", "Related party transactions, if any"];
      NUMLIST(out, std.concat(own, ["Any other business with the permission of the Chairman"]), "");
      out.push(S(), ...SGN(v));
      return out;
    },
  },
  {
    id: "egmmin", group: "Meetings", title: "Minutes of Extraordinary General Meeting", short: "Minutes with attendance, quorum and each resolution as put to the members.",
    basis: "Section 118; Secretarial Standard 2 (general meetings)",
    fields: [{ id: "egmdate", label: "Date of meeting", type: "date" }, { id: "mtime", label: "Time", ph: "11:00 a.m." }, { id: "place", label: "Place", ph: "Registered office" }, { id: "ndate", label: "Date of the notice", type: "date" }, { id: "chair", label: "Chairman of the meeting" },
      { id: "nmem", label: "Members present in person", type: "number" }, { id: "nprox", label: "Members present by proxy or representative", type: "number" },
      { id: "res", label: "Resolutions (one per line: text | Ordinary or Special)", type: "textarea" }, { id: "mode", label: "How the vote was taken", type: "select", opts: [["a show of hands", "Show of hands"], ["a poll", "Poll"], ["remote e-voting and a poll", "E-voting and poll"]] }, { id: "mend", label: "Time the meeting ended", ph: "12:00 noon" }, { id: "sdate", label: "Date of signing", type: "date" }],
    notes: ["Minutes are signed by the Chairman within thirty days of the meeting (Secretarial Standard 2).", "File Form MGT-14 within thirty days for each special resolution."],
    build(v) {
      const out = header(v);
      out.push(T(`MINUTES OF THE EXTRAORDINARY GENERAL MEETING HELD ON ${vd(v, "egmdate", "date").toUpperCase()}${v.mtime ? " AT " + v.mtime.toUpperCase() : ""} AT ${PLACE(v).toUpperCase()}`));
      out.push(P(`**Present:** ${v.nmem ? inrFmt(v.nmem) : "[number]"} members in person and ${v.nprox ? inrFmt(v.nprox) : "[number]"} by proxy or representative, as per the attendance register, and the directors and officers named in it.`));
      out.push(P(`${va(v, "chair", "name")} took the chair. The Chairman noted that the requisite quorum was present, called the meeting to order, and, with the consent of the members, took the notice dated ${vd(v, "ndate", "date")} convening the meeting as read.`));
      LN("res")(v).forEach((x, i) => { const a = x.split("|"), k = /special/i.test(a[1] || "") ? "Special" : "Ordinary"; out.push(P(`**${i + 1}.** The Chairman placed the following resolution before the members as ${k === "Special" ? "a Special" : "an Ordinary"} Resolution:`), P(a[0].trim().replace(/^RESOLVED THAT/i, "**RESOLVED THAT**")), P(`The resolution was put to vote by ${va(v, "mode", "mode")} and, the requisite majority having voted in favour, the Chairman declared it **carried** as ${k === "Special" ? "a Special" : "an Ordinary"} Resolution.`)); });
      out.push(P(`There being no other business, the meeting ended${v.mend ? " at " + v.mend : ""} with a vote of thanks to the Chair.`), S(), L("Date: " + vd(v, "sdate", "date")), L("Place: " + PLACE(v)), R("**" + va(v, "chair", "Chairman") + "**"), R("Chairman"));
      return out;
    },
  },
  {
    id: "pbnotice", group: "Postal ballot", title: "Postal ballot notice and ballot form", short: "Notice, explanatory statement and the postal ballot form, with voting instructions.",
    basis: "Section 110; Rules 20 and 22, Companies (Management and Administration) Rules, 2014; Secretarial Standard 2; Regulation 44 (listed companies)",
    fields: [{ id: "ndate", label: "Date of notice", type: "date" }, { id: "cutoff", label: "Cut-off date for members entitled to vote", type: "date" }, { id: "vstart", label: "Voting starts on", type: "date" }, { id: "vend", label: "Voting ends on", type: "date" },
      { id: "res", label: "Resolutions (one per line: text | Ordinary or Special)", type: "textarea" }, { id: "expl", label: "Explanatory statement (one paragraph per line)", type: "textarea" }, { id: "scr", label: "Scrutinizer's name" }, { id: "scrno", label: "Scrutinizer's membership or CP number" },
      { id: "agency", label: "E-voting agency or platform", ph: "Name of the agency" }, { id: "email", label: "E-mail for the return of ballot forms and queries" }].concat(SG),
    notes: ["Send the notice to every member as on the cut-off date, by e-mail where the company has the address, and give at least thirty days to vote (Secretarial Standard 2).", "Publish an advertisement about the postal ballot in a newspaper in English and in the local language (Rule 22(18)).", "The results go to the stock exchanges and on the website within two working days of the end of voting for a listed company. Use the MGT-13 scrutinizer's report generator."],
    build(v) {
      const out = header(v), res = LN("res")(v).map(x => x.split("|").map(y => y.trim()));
      out.push(T("POSTAL BALLOT NOTICE"), C("[Pursuant to section 110 of the Companies Act, 2013 read with rule 22 of the Companies (Management and Administration) Rules, 2014]"), R("Date: " + vd(v, "ndate", "date")));
      out.push(P(`**NOTICE** is hereby given to the members of **${CO(v)}** that the resolutions set out below are proposed to be passed by postal ballot, including voting by electronic means, under Section 110 and other applicable provisions of the Companies Act, 2013 and the Rules made under it. The explanatory statement under Section 102 is annexed.`));
      (res.length ? res : [["[resolution text]", "Ordinary"]]).forEach((r, i) => { const k = /special/i.test(r[1] || "") ? "Special" : "Ordinary"; out.push(P(`**Item ${i + 1}: ${k} Resolution**`), P(r[0].replace(/^RESOLVED THAT/i, "**RESOLVED THAT**"))); });
      out.push(P("**EXPLANATORY STATEMENT UNDER SECTION 102**")); const ex = LN("expl")(v); (ex.length ? ex : [GAP]).forEach(x => out.push(P(x)));
      out.push(P("**Notes**"));
      ["The notice is sent to the members whose names are on the register of members, or the list of beneficial owners, as on " + vd(v, "cutoff", "cut-off date") + ". A person who is not a member on that date should treat this notice as for information only.", `Voting rights are in proportion to the paid-up equity share capital held on the cut-off date.`, `Voting opens on ${vd(v, "vstart", "start date")} and closes on ${vd(v, "vend", "end date")} (5.00 p.m.). Remote e-voting through ${va(v, "agency", "agency")} is available for that period and is not allowed afterwards.`, `A member may instead complete the attached postal ballot form and send it in the enclosed envelope, or by e-mail to ${va(v, "email", "e-mail")}, so that it reaches the scrutinizer by the close of voting.`, `The Board has appointed ${va(v, "scr", "name")} (${va(v, "scrno", "membership or CP number")}) as the scrutinizer for a fair and transparent process.`, "The results will be announced within the time the law allows and posted on the website of the Company and the agency, and, for a listed company, sent to the stock exchanges."].forEach((x, i) => out.push(I(`${i + 1}. ${x}`)));
      out.push(S(), L("By order of the Board"), ...SGN(v).slice(1));
      out.push(PB(), C("**POSTAL BALLOT FORM**"), ...header(v));
      out.push(TB([["Particulars", "Details"], ["1. Name and registered address of the sole or first-named member", ""], ["2. Name(s) of the joint holder(s), if any", ""], ["3. Registered folio number or DP ID and Client ID", ""], ["4. Number of shares held", ""]], [4500, 4500]));
      out.push(P(`I/We hereby exercise my/our vote in respect of the resolutions set out in the notice dated ${vd(v, "ndate", "date")}, by recording my/our assent or dissent in the following manner:`));
      out.push(TB([["Item no.", "Brief description of the resolution", "No. of shares", "I/We assent (For)", "I/We dissent (Against)"]].concat((res.length ? res : [["[resolution]", "Ordinary"]]).map((r, i) => [String(i + 1), r[0].replace(/^RESOLVED THAT\s*/i, "").slice(0, 110) + (r[0].length > 120 ? "..." : ""), "", "", ""])), [800, 4000, 1300, 1450, 1450]));
      out.push(L("Place: ____________"), L("Date: ____________"), R("Signature of the member: ____________________"), S(), P("**Instructions:** Tick (✓) in the For or Against column against each item. A form that is incomplete, unsigned, defaced, or that marks both columns for an item is invalid for that item. Signatures must match the specimen on record. The form must reach the scrutinizer by the close of voting."));
      return out;
    },
  },
  {
    id: "mgt12", group: "Prescribed formats", title: "Form MGT-12: Polling paper", short: "The polling paper in the layout of Form MGT-12.",
    basis: "Section 109; Rule 21, Companies (Management and Administration) Rules, 2014",
    fields: [{ id: "res", label: "Resolutions put to the poll (one per line)", type: "textarea" }],
    notes: ["This follows the layout of Form MGT-12 as set out in the Rules. Compare it with the current form on the MCA website before use.", "One polling paper goes to each member who votes on the poll."],
    build(v) {
      const out = [C("**Form No. MGT-12**"), C("**POLLING PAPER**"), C("[Pursuant to section 109 of the Companies Act, 2013 and rule 21 of the Companies (Management and Administration) Rules, 2014]"), S(), ...header(v)];
      out.push(C("**BALLOT PAPER**"));
      out.push(TB([["S. No.", "Particulars", "Details"], ["1.", "Name of the first named shareholder (in block letters)", ""], ["2.", "Postal address", ""], ["3.", "Registered folio number / *Client ID number (*applicable to investors holding shares in dematerialised form)", ""], ["4.", "Class of share", ""]], [800, 5200, 3000]));
      out.push(P("I hereby exercise my vote in respect of the Ordinary / Special resolution enumerated below by recording my assent or dissent to the said resolution in the following manner:"));
      const r = LN("res")(v);
      out.push(TB([["No.", "Item no.", "No. of shares held by me", "I assent to the resolution", "I dissent from the resolution"]].concat((r.length ? r : ["[item]"]).map((x, i) => [String(i + 1), x.slice(0, 90), "", "", ""])), [700, 3900, 1500, 1450, 1450]));
      out.push(L("Place: ____________"), L("Date: ____________"), R("Signature of the shareholder: ____________________"));
      return out;
    },
  },
  {
    id: "mgt13", group: "Prescribed formats", title: "Form MGT-13: Report of scrutinizer", short: "The scrutinizer's report on a poll or e-voting, in the layout of Form MGT-13.",
    basis: "Sections 108 and 109; Rules 20 and 21, Companies (Management and Administration) Rules, 2014",
    fields: [{ id: "kind", label: "Meeting", type: "select", opts: [["Annual General Meeting", "Annual General Meeting"], ["Extraordinary General Meeting", "Extraordinary General Meeting"], ["postal ballot", "Postal ballot"]] }, { id: "mdate", label: "Date of meeting or end of voting", type: "date" }, { id: "res", label: "Resolutions (one per line: text | Ordinary or Special)", type: "textarea" },
      { id: "scr", label: "Scrutinizer's name" }, { id: "scraddr", label: "Scrutinizer's address", type: "textarea" }, { id: "scrno", label: "Membership or CP number" }, { id: "chair", label: "Chairman of the meeting" }],
    notes: ["This follows the layout of Form MGT-13 as set out in the Rules: the same headings, the same order of paragraphs and the same result tables. Compare it with the current form on the MCA website before use.", "The report goes to the Chairman within the time the Rules allow after the voting ends (three working days for remote e-voting).", "The scrutinizer writes in the numbers. Nothing here is pre-filled."],
    build(v) {
      const out = [C("**Form No. MGT-13**"), C("**REPORT OF SCRUTINIZER(S)**"), C("[Pursuant to section 109 of the Companies Act, 2013 and rule 21 of the Companies (Management and Administration) Rules, 2014]"), S()];
      out.push(L("To,"), L("The Chairman of the meeting,"), L(`${CO(v)}`), S());
      out.push(P(`Sub.: Report on the voting by ${/postal/i.test(v.kind || "") ? "postal ballot and remote e-voting" : "poll and remote e-voting"} at the ${va(v, "kind", "meeting")} of **${CO(v)}** held on ${vd(v, "mdate", "date")}.`));
      out.push(P(`I/We, ${va(v, "scr", "name")}, appointed as scrutinizer by the Chairman for the purpose of scrutinising the voting process at the above meeting, report as follows:`));
      ["The ballot boxes, in which the poll papers were kept, were opened in my/our presence, found to be empty, and locked and sealed before the poll started. They were again locked and sealed in my/our presence after the poll closed.", "The poll papers received from the members, and the votes cast by remote e-voting, were reconciled with the register of members and the records of the Company and of the Registrar and Share Transfer Agent, and with the authorisations and proxies lodged with the Company.", "The poll papers that were incomplete, defective, unsigned or otherwise invalid are separately listed and were not counted. The number of such papers is: ____________.", "The result of the voting on each resolution is given below."].forEach((x, i) => out.push(I(`${i + 1}. ${x}`)));
      const r = LN("res")(v).map(x => x.split("|").map(y => y.trim()));
      (r.length ? r : [["[resolution]", "Ordinary"]]).forEach((x, i) => {
        out.push(L(`**Resolution no. ${i + 1}: ${/special/i.test(x[1] || "") ? "Special" : "Ordinary"} Resolution** ${x[0].slice(0, 120)}`));
        out.push(TB([["Particulars", "Total number of members / votes", "Votes in favour", "Votes against", "Invalid votes"], ["Number of members who voted", "", "", "", ""], ["Number of valid votes (shares)", "", "", "", ""], ["Percentage of the total valid votes", "100%", "", "", ""]], [3000, 1800, 1400, 1400, 1400]));
      });
      ["A compact disc (or a soft copy) listing the shareholders who voted For, who voted Against, and whose votes were invalid, for each resolution, is enclosed.", "The poll papers, the other relevant records and the e-voting data were sealed and handed over to the Company Secretary, or to the director authorised by the Board, for safe custody."].forEach((x, i) => out.push(I(`${i + 5}. ${x}`)));
      out.push(S(), L("Place: ____________"), L("Date: ____________"), R("Signature of the Scrutinizer: ____________________"), R(`**${va(v, "scr", "Name")}**`), R(va(v, "scrno", "Membership or CP number")), R((v.scraddr || "").replace(/\n+/g, ", ") || "[address]"), S(), L("Received and countersigned by the Chairman:"), R(`**${va(v, "chair", "Chairman")}**`));
      return out;
    },
  },
  {
    id: "mr3", group: "Prescribed formats", title: "Form MR-3: Secretarial audit report", short: "The secretarial audit report in the layout of Form MR-3, with Annexure A.",
    basis: "Section 204(1), Companies Act, 2013; Rule 9, Companies (Appointment and Remuneration of Managerial Personnel) Rules, 2014",
    fields: [{ id: "fyend", label: "Financial year ended on", type: "date" }, { id: "listed", label: "Is the company listed?", type: "select", opts: [["yes", "Yes, listed"], ["no", "No, unlisted"]] }, { id: "other", label: "Other laws specifically applicable to the company (one per line)", type: "textarea" },
      { id: "obs", label: "Observations on the compliances above (one per line, or leave blank for none)", type: "textarea" }, { id: "events", label: "Specific events during the period (one per line: public or preferential issue, buy-back, section 180 decisions, merger, foreign collaboration)", type: "textarea" },
      { id: "place", label: "Place" }, { id: "date", label: "Date of report", type: "date" }, { id: "csname", label: "Name of the Company Secretary in practice or firm" }, { id: "memno", label: "ACS/FCS number" }, { id: "cpno", label: "CP number" }, { id: "udin", label: "UDIN" }],
    notes: ["This follows the layout of Form MR-3 as set out in the Rules: the same opening, the same list of laws, the same 'I further report' paragraphs and the same Annexure A. Compare it with the current form on the MCA website.", "Listed companies must also follow the SEBI requirements for the secretarial audit report, and for the annual secretarial compliance report under Regulation 24A.", "Remove from the list any law that does not apply to the company."],
    build(v) {
      const co = CO(v), fy = vd(v, "fyend", "date"), listed = v.listed !== "no";
      const out = [C("**Form No. MR-3**"), C("**SECRETARIAL AUDIT REPORT**"), C(`FOR THE FINANCIAL YEAR ENDED ${fy.toUpperCase()}`), C("[Pursuant to section 204(1) of the Companies Act, 2013 and rule No. 9 of the Companies (Appointment and Remuneration Personnel) Rules, 2014]"), S(), L("To,"), L("The Members,"), L("**" + co + "**"), ...(v.ro ? [L(v.ro.replace(/\n+/g, ", "))] : []), S()];
      out.push(P(`I have conducted the secretarial audit of the compliance of applicable statutory provisions and the adherence to good corporate practices by **${co}** (hereinafter called the Company). Secretarial Audit was conducted in a manner that provided me a reasonable basis for evaluating the corporate conducts/statutory compliances and expressing my opinion thereon.`));
      out.push(P(`Based on my verification of the Company's books, papers, minute books, forms and returns filed and other records maintained by the Company and also the information provided by the Company, its officers, agents and authorised representatives during the conduct of secretarial audit, I hereby report that in my opinion, the Company has, during the audit period covering the financial year ended on ${fy} complied with the statutory provisions listed hereunder and also that the Company has proper Board-processes and compliance-mechanism in place to the extent, in the manner and subject to the reporting made hereinafter:`));
      out.push(P(`I have examined the books, papers, minute books, forms and returns filed and other records maintained by ${co} for the financial year ended on ${fy} according to the provisions of:`));
      const laws = ["The Companies Act, 2013 (the Act) and the rules made thereunder;", "The Securities Contracts (Regulation) Act, 1956 ('SCRA') and the rules made thereunder;", "The Depositories Act, 1996 and the Regulations and Bye-laws framed thereunder;", "Foreign Exchange Management Act, 1999 and the rules and regulations made thereunder to the extent of Foreign Direct Investment, Overseas Direct Investment and External Commercial Borrowings;"];
      if (listed) laws.push("The following Regulations and Guidelines prescribed under the Securities and Exchange Board of India Act, 1992 ('SEBI Act'):", "(a) The Securities and Exchange Board of India (Substantial Acquisition of Shares and Takeovers) Regulations, 2011;", "(b) The Securities and Exchange Board of India (Prohibition of Insider Trading) Regulations, 2015;", "(c) The Securities and Exchange Board of India (Issue of Capital and Disclosure Requirements) Regulations, 2018;", "(d) The Securities and Exchange Board of India (Share Based Employee Benefits and Sweat Equity) Regulations, 2021;", "(e) The Securities and Exchange Board of India (Issue and Listing of Non-Convertible Securities) Regulations, 2021;", "(f) The Securities and Exchange Board of India (Registrars to an Issue and Share Transfer Agents) Regulations, 1993;", "(g) The Securities and Exchange Board of India (Delisting of Equity Shares) Regulations, 2021;", "(h) The Securities and Exchange Board of India (Buyback of Securities) Regulations, 2018;");
      const RM = ["i", "ii", "iii", "iv", "v", "vi", "vii"]; let ln = 0;
      laws.forEach(x => { if (/^\(/.test(x)) out.push(I(x, 2)); else out.push(I(`(${RM[ln++]}) ${x}`)); });
      const oth = LN("other")(v); out.push(I(`(${RM[ln++]}) The other laws specifically applicable to the Company: ${oth.length ? "" : "[list]"}`)); oth.forEach(x => out.push(I("- " + x, 2)));
      out.push(P("I have also examined compliance with the applicable clauses of the following:"), I("(i) Secretarial Standards issued by The Institute of Company Secretaries of India."), I("(ii) " + (listed ? "The Securities and Exchange Board of India (Listing Obligations and Disclosure Requirements) Regulations, 2015 and the Listing Agreements entered into by the Company with the Stock Exchanges." : "The Listing Agreements, if any (not applicable to an unlisted company).")));
      const ob = LN("obs")(v);
      out.push(P("During the period under review the Company has complied with the provisions of the Act, Rules, Regulations, Guidelines, Standards, etc. mentioned above" + (ob.length ? " subject to the following observations:" : ".")));
      ob.forEach((x, i) => out.push(I(`${i + 1}. ${x}`)));
      out.push(P("I further report that:"));
      out.push(P("The Board of Directors of the Company is duly constituted with proper balance of Executive Directors, Non-Executive Directors and Independent Directors. The changes in the composition of the Board of Directors that took place during the period under review were carried out in compliance with the provisions of the Act."));
      out.push(P("Adequate notice is given to all directors to schedule the Board Meetings, agenda and detailed notes on agenda were sent at least seven days in advance, and a system exists for seeking and obtaining further information and clarifications on the agenda items before the meeting and for meaningful participation at the meeting."));
      out.push(P("Majority decision is carried through while the dissenting members' views are captured and recorded as part of the minutes."));
      out.push(P("I further report that there are adequate systems and processes in the Company commensurate with the size and operations of the Company to monitor and ensure compliance with applicable laws, rules, regulations and guidelines."));
      const ev = LN("events")(v);
      out.push(P("I further report that during the audit period the Company has had the following specific events/actions having a major bearing on the Company's affairs in pursuance of the above referred laws, rules, regulations, guidelines, standards, etc.:"));
      (ev.length ? ev : ["Public/Rights/preferential issue of shares/debentures/sweat equity, etc.: [details or nil]", "Redemption/buy-back of securities: [details or nil]", "Major decisions taken by the members in pursuance to Section 180 of the Companies Act, 2013: [details or nil]", "Merger/amalgamation/reconstruction, etc.: [details or nil]", "Foreign technical collaborations: [details or nil]"]).forEach((x, i) => out.push(I(`(${["i", "ii", "iii", "iv", "v", "vi", "vii", "viii"][i] || i + 1}) ${x}`)));
      out.push(S(), L("Place: " + va(v, "place", "place")), L("Date: " + vd(v, "date", "date")), R("**" + va(v, "csname", "Name of the Company Secretary in practice") + "**"), R("ACS/FCS No.: " + va(v, "memno", "number")), R("C.P. No.: " + va(v, "cpno", "number")), R("UDIN: " + va(v, "udin", "UDIN")));
      out.push(P("**Note:** This report is to be read with my letter of even date which is annexed as Annexure A and forms an integral part of this report."));
      out.push(PB(), C("**Annexure A**"), L("To,"), L("The Members,"), L("**" + co + "**"), S(), P("My report of even date is to be read along with this letter."));
      ["Maintenance of secretarial record is the responsibility of the management of the Company. My responsibility is to express an opinion on these secretarial records based on my audit.", "I have followed the audit practices and processes as were appropriate to obtain reasonable assurance about the correctness of the contents of the secretarial records. The verification was done on test basis to ensure that correct facts are reflected in secretarial records. I believe that the processes and practices, I followed provide a reasonable basis for my opinion.", "I have not verified the correctness and appropriateness of financial records and books of accounts of the Company.", "Wherever required, I have obtained the management representation about the compliance of laws, rules and regulations and happening of events, etc.", "The compliance of the provisions of corporate and other applicable laws, rules, regulations, standards is the responsibility of management. My examination was limited to the verification of procedure on test basis.", "The Secretarial Audit report is neither an assurance as to the future viability of the Company nor of the efficacy or effectiveness with which the management has conducted the affairs of the Company."].forEach((x, i) => out.push(P(`${i + 1}. ${x}`)));
      out.push(S(), L("Place: " + va(v, "place", "place")), L("Date: " + vd(v, "date", "date")), R("**" + va(v, "csname", "Name") + "**"), R("ACS/FCS No.: " + va(v, "memno", "number")), R("C.P. No.: " + va(v, "cpno", "number")));
      return out;
    },
  },
  {
    id: "rptpol", group: "Policies", title: "Related party transactions policy", short: "Policy on identifying, approving and reporting related party transactions.",
    basis: "Section 188, Companies Act, 2013; Rule 15; Regulation 23, SEBI (LODR) Regulations, 2015",
    fields: POL_FIELDS.slice(0, 2).concat([{ id: "limit", label: "Value above which the Board's approval is always needed (Rs.)", type: "number" }, { id: "omni", label: "Period for which the Audit Committee's omnibus approval is valid", ph: "one year" }, { id: "contact", label: "Name of the contact person", ph: "Company Secretary" }]),
    notes: ["Listed companies must have this policy and put it on the website (Regulation 23). Others use it as good practice.", "Check the current thresholds in Rule 15 and in Regulation 23 before you fix the limits."],
    build(v) {
      const o = polHead(v, "POLICY ON RELATED PARTY TRANSACTIONS", "Under Section 188 of the Companies Act, 2013 and Regulation 23 of the SEBI (LODR) Regulations, 2015"); let n = 1;
      clause(o, n++, "Purpose", `${CO(v)} wants every transaction with a related party to be fair, transparent and in the interest of the Company and its members. This policy sets out how such transactions are identified, approved and disclosed.`);
      clause(o, n++, "Definitions", "'Related party' and 'related party transaction' have the meanings in Section 2(76) of the Companies Act, 2013 and, for a listed company, in Regulation 2(1)(zb) and (zc) of the SEBI (LODR) Regulations, 2015. 'Arm's length' means a transaction between two unrelated parties that is carried out in their own interests.");
      clause(o, n++, "Identification", "Every director and key managerial personnel gives the Company Secretary a list of the persons and entities in which they or their relatives are interested, at the first Board meeting of each year and when anything changes. The Company keeps a list of related parties and updates it.");
      clause(o, n++, "Approval of the Audit Committee", `Every related party transaction needs the prior approval of the Audit Committee. The Committee may give an omnibus approval for transactions that are repetitive in nature and in the ordinary course of business, valid for ${va(v, "omni", "period")}, as the law allows. Only independent directors on the Committee approve the transactions.`);
      clause(o, n++, "Approval of the Board and the members", `A transaction that is not in the ordinary course of business or not at arm's length, or that is above ${v.limit ? "Rs. " + inrFmt(v.limit) : "[limit]"} or the thresholds in the law, goes to the Board. A transaction above the thresholds in Rule 15, or the material limit in Regulation 23, also needs the approval of the members by a resolution on which related parties do not vote.`);
      clause(o, n++, "Transactions not approved in advance", "The Audit Committee may ratify a transaction that was entered into without prior approval only in the cases and within the time the law allows. The Company records the reasons.");
      clause(o, n++, "Disclosure", "The Company discloses related party transactions in Form AOC-2 attached to the Board's report, in the financial statements as required by the accounting standards, and, for a listed company, to the stock exchanges every half year and on its website.");
      clause(o, n++, "Review", `The Audit Committee reviews this policy at least once in three years. The contact person is ${va(v, "contact", "name")}.`);
      return o;
    },
  },
  {
    id: "conduct", group: "Policies", title: "Code of conduct for directors and senior management", short: "A code with the duties of directors under Section 166 and the code for independent directors.",
    basis: "Section 166 and Schedule IV, Companies Act, 2013; Regulation 17(5), SEBI (LODR) Regulations, 2015",
    fields: POL_FIELDS,
    notes: ["Listed companies must have a code of conduct for the Board and senior management, put it on the website, and take an annual affirmation of compliance (Regulation 26(3)).", "Independent directors must follow the code in Schedule IV."],
    build(v) {
      const o = polHead(v, "CODE OF CONDUCT FOR DIRECTORS AND SENIOR MANAGEMENT", "Under Section 166 and Schedule IV of the Companies Act, 2013 and Regulation 17(5) of the SEBI (LODR) Regulations, 2015"); let n = 1;
      clause(o, n++, "Who it applies to", `All the directors and the senior management of ${CO(v)}. Independent directors also follow the Code for Independent Directors in Schedule IV.`);
      clause(o, n++, "Act in good faith", "A director acts in accordance with the articles of the Company, in good faith, to promote the objects of the Company for the benefit of its members as a whole, and in the best interests of the Company, its employees, the members, the community and the protection of the environment (Section 166(2)).");
      clause(o, n++, "Care and independent judgement", "A director exercises duties with due and reasonable care, skill and diligence and exercises independent judgement (Section 166(3)).");
      clause(o, n++, "Conflicts of interest", "A director must not be in a situation in which he or she has, or can have, a direct or indirect interest that conflicts, or may conflict, with the interest of the Company, and must disclose any interest under Section 184 (Section 166(4)).");
      clause(o, n++, "No undue gain", "A director must not achieve or attempt to achieve any undue gain or advantage, either to himself or herself or to relatives, partners or associates, and must pay back any such gain to the Company if it is made (Section 166(5)).");
      clause(o, n++, "No assignment of office", "A director must not assign his or her office, and any assignment is void (Section 166(6)).");
      clause(o, n++, "Confidentiality and insider information", "Directors and senior management keep the Company's confidential information safe, do not use it for personal gain, and follow the Company's code to regulate and monitor trading in its securities.");
      clause(o, n++, "Compliance with laws", "Directors and senior management follow the laws that apply to the Company, and report any breach that they become aware of.");
      clause(o, n++, "Additional duties of independent directors", "In addition, an independent director follows Schedule IV: he or she helps bring an independent judgement to Board deliberations, safeguards the interests of all the stakeholders, particularly the minority shareholders, attends the meetings, and does not seek to benefit from the Company other than as the law allows.");
      clause(o, n++, "Annual affirmation", `Every director and every member of the senior management confirms in writing each year that he or she has complied with this Code. The confirmations go to ${va(v, "contact", "the Company Secretary")} at ${va(v, "email", "e-mail")}.`);
      return o;
    },
  },
  {
    id: "matpol", group: "Policies", title: "Policy for determining materiality of events", short: "Policy for deciding which events and information the company must disclose to the exchanges.",
    basis: "Regulation 30, SEBI (LODR) Regulations, 2015",
    fields: POL_FIELDS.slice(0, 2).concat([{ id: "kmp", label: "Key managerial personnel authorised to decide on disclosure (one per line)", type: "textarea" }, { id: "contact", label: "Name of the contact person", ph: "Company Secretary" }, { id: "email", label: "E-mail of the contact person" }]),
    notes: ["Only listed companies need this policy (Regulation 30(4)(ii)). Put it on the website.", "Check the current tests of materiality in Regulation 30(4)(i), and the list of events in Part A of Schedule III, before you adopt this."],
    build(v) {
      const o = polHead(v, "POLICY FOR DETERMINATION OF MATERIALITY OF EVENTS OR INFORMATION", "Under Regulation 30 of the SEBI (LODR) Regulations, 2015"); let n = 1;
      clause(o, n++, "Purpose", `${CO(v)} wants to give the stock exchanges and the market timely and accurate information about events that matter. This policy sets the criteria for deciding when an event or information is material.`);
      clause(o, n++, "Events that are always disclosed", "The events in Para A of Part A of Schedule III to the Regulations, such as an acquisition, a change in the capital, a scheme of arrangement, a change in the directors or key managerial personnel, a fraud or default, and the outcome of the Board meeting, are disclosed without a test of materiality, within the time the Regulations give.");
      clause(o, n++, "Events disclosed if material", "The events in Para B of Part A of Schedule III are disclosed when they are material under the tests in Regulation 30(4)(i): the omission of the event is likely to result in a discontinuity or alteration of an event already in the public domain, or in the opinion of the Board the event is likely to affect the market, or the value of the event or information crosses the lower of two per cent of turnover, two per cent of net worth, or five per cent of the average of the absolute value of profit or loss after tax of the last three financial years, as stated in the Regulations.");
      clause(o, n++, "Other events", "Any other information that is price sensitive, or that the Board thinks it would be fair to share, is also disclosed.");
      clause(o, n++, "Who decides", "The following officers decide whether an event is material and make the disclosure, and their contact details are on the website:"); NUMLIST(o, LN("kmp")(v), "[officer]");
      clause(o, n++, "Time of disclosure", "The Company discloses as soon as it can, and in any case within thirty minutes of the close of the Board meeting for its outcome, within twelve hours for events that happen outside the Company, and within twenty-four hours for the others, as Regulation 30(6) says.");
      clause(o, n++, "Website and records", `The Company keeps the disclosures on its website for at least five years and then as its archival policy says. The contact person is ${va(v, "contact", "name")} (${va(v, "email", "e-mail")}).`);
      clause(o, n++, "Review", "The Board reviews this policy at least once a year and when the Regulations change.");
      return o;
    },
  },
);
/* ---------- batch 4: secretarial letters (own wording) ---------- */
const LF = [{ id: "ldate", label: "Date of letter", type: "date" }, { id: "ref", label: "Reference number (optional)" }];
const LSG = [{ id: "sname", label: "Name of the signatory" }, { id: "sdes", label: "Designation", ph: "Company Secretary" }];
function letter(v, to, subject, paras, closing) {
  const o = header(v);
  o.push(L("Ref: " + ((v.ref || "").trim() || "[number]")), R("Date: " + vd(v, "ldate", "date")), S(), L("To,"));
  to.forEach(x => o.push(L(x))); o.push(S(), P("**Subject: " + subject + "**"), S(), L("Dear Sir / Madam,"), S());
  paras.forEach(x => o.push(P(x))); o.push(S(), L(closing || "Yours faithfully,"), R("For **" + CO(v) + "**"), S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")));
  return o;
}
const ROC = v => ["The Registrar of Companies,", va(v, "roc", "ROC office, city"), ...(v.rocaddr ? [v.rocaddr.replace(/\n+/g, ", ")] : [])];
const ROCF = [{ id: "roc", label: "Registrar of Companies office", ph: "RoC Mumbai" }];
const dirName = v => va(v, "dname", "name of director");

GENS.push(
  {
    id: "l-rocreply", group: "Letters to the Registrar and exchanges", title: "Reply to a notice from the Registrar", short: "A reply to a notice or query from the ROC, point by point.",
    basis: "Sections 206, 207 and 454, Companies Act, 2013 (as the notice states)",
    fields: ROCF.concat([{ id: "nno", label: "Notice number" }, { id: "ndate", label: "Date of notice", type: "date" }, { id: "nsec", label: "Section the notice is issued under" }, { id: "points", label: "Your reply to each point in the notice (one per line)", type: "textarea" }, { id: "encl", label: "Documents enclosed (one per line)", type: "textarea" }], LF, LSG),
    notes: ["Reply within the time the notice gives. If you need more time, write for an extension before the date.", "Send the reply through the MCA portal if the notice asks for it."],
    build(v) {
      const pts = LN("points")(v), en = LN("encl")(v);
      const o = letter(v, ROC(v), `Reply to notice no. ${va(v, "nno", "number")} dated ${vd(v, "ndate", "date")} issued under Section ${va(v, "nsec", "section")} of the Companies Act, 2013 to ${CO(v)}${v.cin ? " (CIN: " + v.cin + ")" : ""}`,
        [`We refer to your notice mentioned above. We thank you for giving us the opportunity to explain, and we submit our reply to each point in it.`]);
      const i = o.findIndex(b => b.k === "L" && b.t === "Yours faithfully,");
      const add = []; (pts.length ? pts : ["[reply to the point]"]).forEach((x, k) => add.push(I(`${k + 1}. ${x}`)));
      add.push(P("We request you to take our reply on record and to close the matter. We are ready to give any further information or documents that you may need."));
      if (en.length) { add.push(P("**Enclosures**")); en.forEach((x, k) => add.push(I(`${k + 1}. ${x}`))); }
      o.splice(i - 1, 0, ...add); return o;
    },
  },
  {
    id: "l-rocrect", group: "Letters to the Registrar and exchanges", title: "Request to the Registrar to correct a record", short: "A request to correct an error in the company's record, or in a filed form.",
    basis: "Section 403 and Section 7(5)/(6) as applicable; the MCA's procedure for rectification",
    fields: ROCF.concat([{ id: "what", label: "What is wrong (form, SRN and the error)", type: "textarea" }, { id: "right", label: "What it should say", type: "textarea" }, { id: "why", label: "How the error happened", type: "textarea" }, { id: "encl", label: "Documents enclosed (one per line)", type: "textarea" }], LF, LSG),
    notes: ["Many corrections are made by filing the form again or a specific rectification form on the MCA portal. Use this letter with that filing or where the portal needs a written request."],
    build(v) {
      const en = LN("encl")(v);
      const o = letter(v, ROC(v), `Request to correct the record of ${CO(v)}${v.cin ? " (CIN: " + v.cin + ")" : ""}`, [`We write to bring to your notice an error in the record of the Company. ${va(v, "what", "describe the error")}`, `The correct position is: ${va(v, "right", "correct details")}`, `The error happened because ${va(v, "why", "explain how it happened")}. It was not intended to mislead anyone.`, "We request you to correct the record. We attach the documents that support the correction." + (en.length ? " The documents are: " + en.join("; ") + "." : "")]);
      return o;
    },
  },
  {
    id: "l-bm29", group: "Letters to the Registrar and exchanges", title: "Intimation to the stock exchange of a Board meeting", short: "Prior intimation of a Board meeting and its purpose under Regulation 29.",
    basis: "Regulation 29, SEBI (LODR) Regulations, 2015",
    fields: [{ id: "exch", label: "Stock exchange(s)", ph: "BSE Limited / National Stock Exchange of India Limited" }, { id: "scrip", label: "Scrip code and symbol" }, { id: "mdate", label: "Date of Board meeting", type: "date" }, { id: "purpose", label: "Purpose (one per line)", type: "textarea", ph: "To consider and approve the financial results for the quarter ended 30 September" }, { id: "window", label: "Trading window closed from", type: "date" }].concat(LF, LSG),
    notes: ["Give at least two clear working days' notice, excluding the date of the intimation and of the meeting, for the items in Regulation 29(1) (check the current rule).", "For a meeting on financial results, also check the notice period for the trading window under the insider trading regulations."],
    build(v) {
      return letter(v, ["The Listing Department,", va(v, "exch", "stock exchange"), ""], `Intimation of a Board meeting under Regulation 29 of the SEBI (LODR) Regulations, 2015 (Scrip code: ${va(v, "scrip", "code and symbol")})`, [`Pursuant to Regulation 29 of the SEBI (Listing Obligations and Disclosure Requirements) Regulations, 2015, we inform you that a meeting of the Board of Directors of the Company will be held on **${vd(v, "mdate", "date")}**, to consider the following:`].concat(LN("purpose")(v).map((x, i) => `${i + 1}. ${x}`)).concat(v.window ? [`The trading window for dealing in the securities of the Company has been closed from ${vd(v, "window", "date")} for the designated persons, as the insider trading code requires, and will open 48 hours after the results are made public.`] : []).concat(["Kindly take this on record."]));
    },
  },
  {
    id: "l-outcome", group: "Letters to the Registrar and exchanges", title: "Outcome of Board meeting to the stock exchange", short: "Intimation of the decisions of a Board meeting under Regulation 30.",
    basis: "Regulation 30 and Part A of Schedule III, SEBI (LODR) Regulations, 2015",
    fields: [{ id: "exch", label: "Stock exchange(s)" }, { id: "scrip", label: "Scrip code and symbol" }, { id: "mdate", label: "Date of Board meeting", type: "date" }, { id: "start", label: "Meeting commenced at", ph: "11:00 a.m." }, { id: "end", label: "Meeting concluded at", ph: "1:30 p.m." }, { id: "dec", label: "Decisions (one per line)", type: "textarea" }].concat(LF, LSG),
    notes: ["File the outcome within thirty minutes of the close of the meeting (Regulation 30(6)). Attach the financial results and the auditor's limited review report if the meeting approved them."],
    build(v) {
      return letter(v, ["The Listing Department,", va(v, "exch", "stock exchange"), ""], `Outcome of the Board meeting held on ${vd(v, "mdate", "date")} (Scrip code: ${va(v, "scrip", "code and symbol")})`, [`Pursuant to Regulations 30 and 33 and other applicable provisions of the SEBI (Listing Obligations and Disclosure Requirements) Regulations, 2015, we inform you that the Board of Directors, at its meeting held today, ${vd(v, "mdate", "date")}, which commenced at ${va(v, "start", "time")} and concluded at ${va(v, "end", "time")}, has considered and approved the following:`].concat(LN("dec")(v).map((x, i) => `${i + 1}. ${x}`)).concat(["Kindly take this on record."]));
    },
  },
  {
    id: "l-auditorint", group: "Letters to directors, auditors and members", title: "Intimation of appointment to the auditor", short: "The letter that tells the auditor of the appointment, under Section 139(1).",
    basis: "Section 139(1) and (2); Rule 4, Companies (Audit and Auditors) Rules, 2014",
    fields: [{ id: "firm", label: "Name of the auditor or firm" }, { id: "faddr", label: "Address of the auditor", type: "textarea" }, { id: "agmno", label: "Which AGM appointed them", ph: "fifth" }, { id: "agmdate", label: "Date of that AGM", type: "date" }, { id: "from", label: "Financial year they start with", ph: "2026-27" }, { id: "to", label: "Appointed until the AGM for the year", ph: "2030-31" }].concat(LF, LSG),
    notes: ["The company must inform the auditor of the appointment, and the auditor must file Form ADT-1 within fifteen days of the meeting. If your auditor is reappointed, say so in place of 'appointed'."],
    build(v) {
      return letter(v, [va(v, "firm", "auditor"), (v.faddr || "[address]").replace(/\n+/g, ", ")], `Intimation of your appointment as statutory auditor of ${CO(v)}`, [`We are pleased to inform you that the members of the Company, at the ${va(v, "agmno", "number")} Annual General Meeting held on ${vd(v, "agmdate", "date")}, have appointed you as the statutory auditor of the Company under Section 139(1) of the Companies Act, 2013, from the financial year ${va(v, "from", "year")}, to hold office until the conclusion of the Annual General Meeting for the financial year ${va(v, "to", "year")}.`, "As Rule 4 requires, we ask you to file Form ADT-1 with the Registrar of Companies within fifteen days of the meeting. Please send us a copy of the filing. We also ask you to confirm that your appointment is in line with Section 141 of the Act.", "We look forward to your work with us."]);
    },
  },
  {
    id: "l-indapp", group: "Letters to directors, auditors and members", title: "Letter of appointment of an independent director", short: "A letter of appointment with the terms Schedule IV requires.",
    basis: "Section 149(4), (6) and (7); Schedule IV (Part VI), Companies Act, 2013; Regulation 17(1A) and 25 (listed companies)",
    fields: [{ id: "dname", label: "Name of the director" }, { id: "daddr", label: "Address", type: "textarea" }, { id: "adate", label: "Date of appointment", type: "date" }, { id: "tenure", label: "Term", ph: "five consecutive years" }, { id: "fee", label: "Sitting fee per meeting (Rs.)", type: "number" }, { id: "commits", label: "Committees (one per line)", type: "textarea" }, { id: "mdate", label: "Date of the Board resolution", type: "date" }].concat(LF, LSG),
    notes: ["Schedule IV requires a formal letter of appointment with these terms, and the company must put the terms on its website. Independent directors are not eligible for stock options and cannot take a commission linked to profits in some cases. Check the current rules."],
    build(v) {
      const c = LN("commits")(v), o = letter(v, [dirName(v), (v.daddr || "[address]").replace(/\n+/g, ", ")], `Appointment as an Independent Director of ${CO(v)}`, [`We are glad to inform you that the Board of Directors, at its meeting held on ${vd(v, "mdate", "date")}, appointed you as a Non-Executive Independent Director of the Company with effect from ${vd(v, "adate", "date")}, for a term of ${va(v, "tenure", "term")}, subject to the approval of the members. This letter sets out the terms of the appointment, as Schedule IV to the Companies Act, 2013 requires.`]);
      const i = o.findIndex(b => b.k === "L" && b.t === "Yours faithfully,"), a = [];
      [["Term", "Your term is as stated above. You are not liable to retire by rotation. A second term needs a special resolution of the members, after the Board has found that you still meet the criteria."], ["Expected role", "You will bring independent judgement to the Board's deliberations on strategy, performance, risk management, resources, key appointments and standards of conduct. You will help the Board protect the interests of all the stakeholders, especially the minority shareholders."], ["Committees", c.length ? "You are a member of the following committees: " + c.join("; ") + "." : "The Board may ask you to serve on its committees."], ["Code of conduct", "You will follow the Code for Independent Directors in Schedule IV, the Company's code of conduct, and the duties of directors in Section 166."], ["Time", "You are expected to attend all the Board and committee meetings and the general meetings, and to give the time the role needs."], ["Fees", `You will get a sitting fee of ${v.fee ? "Rs. " + inrFmt(v.fee) : "[amount]"} for each Board and committee meeting you attend, and reimbursement of your reasonable expenses. You will not get stock options.`], ["Disclosures", "You will give the Company a declaration of independence at the first meeting of each financial year and when anything changes, and a notice of disclosure of interest under Section 184. You will inform the Company at once if you cease to meet the criteria of independence."], ["Insurance and indemnity", "The Company will include you in its directors' and officers' liability insurance, if it has one, as the law allows."], ["Confidentiality and insider information", "You will keep the Company's information confidential and follow its code to regulate and monitor trading in its securities."], ["Termination", "Your office ends if you resign, are removed as the Act allows, or become disqualified. The Company will put the terms of this appointment on its website."]].forEach((x, k) => a.push(P(`**${k + 1}. ${x[0]}.** ${x[1]}`)));
      a.push(P("Please sign and return a copy of this letter to show your acceptance of the terms.")); o.splice(i - 1, 0, ...a);
      o.push(S(), L("I accept the terms of this appointment."), L("Signature: ____________________"), L("Name: " + dirName(v)), L("Date: ____________")); return o;
    },
  },
  {
    id: "l-kmp", group: "Letters to directors, auditors and members", title: "Appointment letter for Company Secretary or CFO", short: "An appointment letter for a key managerial person, with the main terms.",
    basis: "Sections 203 and 2(51), Companies Act, 2013; Rule 8 and 8A, Companies (Appointment and Remuneration of Managerial Personnel) Rules, 2014",
    fields: [{ id: "dname", label: "Name of the appointee" }, { id: "daddr", label: "Address", type: "textarea" }, { id: "post", label: "Post", type: "select", opts: [["Company Secretary", "Company Secretary"], ["Chief Financial Officer", "Chief Financial Officer"], ["Chief Executive Officer", "Chief Executive Officer"], ["Manager", "Manager"]] }, { id: "adate", label: "Date of joining", type: "date" }, { id: "mdate", label: "Date of Board resolution", type: "date" }, { id: "pay", label: "Annual cost to company (Rs.)", type: "number" }, { id: "notice", label: "Notice period", ph: "three months" }, { id: "probation", label: "Probation", ph: "six months" }].concat(LF, LSG),
    notes: ["Key managerial personnel must be appointed by a Board resolution that sets the terms. File Form DIR-12 within thirty days of the appointment.", "A Company Secretary must be a member of the ICSI. Add your standard HR terms to this letter."],
    build(v) {
      const o = letter(v, [dirName(v), (v.daddr || "[address]").replace(/\n+/g, ", ")], `Appointment as ${va(v, "post", "post")}`, [`We are pleased to inform you that the Board of Directors of the Company, at its meeting held on ${vd(v, "mdate", "date")}, has appointed you as the **${va(v, "post", "post")}** of the Company, a key managerial personnel under Section 203 of the Companies Act, 2013, with effect from ${vd(v, "adate", "date")}, on the terms below.`]);
      const i = o.findIndex(b => b.k === "L" && b.t === "Yours faithfully,"), a = [];
      [["Duties", "You will perform the duties that the Companies Act, 2013 and the other laws give to your office, the duties the Board gives you, and the duties in the Company's policies. You will report to the Board."], ["Remuneration", `Your annual cost to company is ${v.pay ? vmoney(v, "pay") : "[amount]"}, with the breakup in the annexure. Deductions are made as the law requires.`], ["Probation", `You are on probation for ${va(v, "probation", "period")}. The Company will confirm you in writing.`], ["Working hours and place", "Your place of work is the registered office of the Company or such other place as the Company may decide. You will give the time that the role needs."], ["Confidentiality and conflicts", "You will keep the Company's information confidential, will not work for another employer without our written consent, and will disclose any interest that conflicts with the Company's."], ["Notice period", `Either side may end the appointment by giving ${va(v, "notice", "notice")} of notice, or pay in lieu of notice. The Company may end the appointment at once for misconduct.`], ["Filing", "The Company will file Form DIR-12 with the Registrar of Companies within thirty days of your appointment."]].forEach((x, k) => a.push(P(`**${k + 1}. ${x[0]}.** ${x[1]}`)));
      a.push(P("Please sign and return a copy of this letter to show your acceptance.")); o.splice(i - 1, 0, ...a);
      o.push(S(), L("I accept the appointment on these terms."), L("Signature: ____________________"), L("Name: " + dirName(v)), L("Date: ____________")); return o;
    },
  },
  {
    id: "l-dresign", group: "Letters to directors, auditors and members", title: "Resignation letter of a director", short: "A director's resignation letter with the DIR-11 reminder, and a reply acknowledging it.",
    basis: "Section 168, Companies Act, 2013; Rule 16, Companies (Appointment and Qualification of Directors) Rules, 2014",
    fields: [{ id: "dname", label: "Name of the director" }, { id: "din", label: "DIN" }, { id: "rdate", label: "Date the resignation takes effect", type: "date" }, { id: "reason", label: "Reason (optional)", ph: "Personal reasons / other commitments" }, { id: "chair", label: "Chairman or person the letter is addressed to", ph: "The Chairman" }].concat(LF),
    notes: ["The resignation takes effect from the date the Company receives the notice, or the date in the notice if it is later. The director may also file Form DIR-11, and the Company must file Form DIR-12 within thirty days.", "A director who resigns may put the reasons in DIR-11. The Company must also state the reasons in the Board's report."],
    build(v) {
      const o = header(v);
      o.push(R("Date: " + vd(v, "ldate", "date")), S(), L("To,"), L(va(v, "chair", "The Chairman")), L("The Board of Directors,"), L(CO(v)), S(), P(`**Subject: Resignation from the office of Director**`), S(), L("Dear Sir,"), S());
      o.push(P(`I, ${dirName(v)} (DIN: ${va(v, "din", "DIN")}), give notice under Section 168 of the Companies Act, 2013 of my resignation from the office of Director of ${CO(v)}, with effect from ${vd(v, "rdate", "date")}${v.reason ? ", for the following reason: " + v.reason : ""}.`), P("I request you to accept my resignation, to record it in the minutes of the next Board meeting, and to file the necessary form with the Registrar of Companies. I will also file Form DIR-11 myself. I thank the Board and the Company for the opportunity to serve."), S(), L("Yours faithfully,"), S(), S(), L("____________________"), L("**" + dirName(v) + "**"), L("DIN: " + va(v, "din", "DIN")));
      o.push(PB(), ...header(v), R("Date: ____________"), S(), L("To,"), L(dirName(v)), S(), P("**Subject: Acknowledgement of your resignation**"), S(), L("Dear Sir,"), S(), P(`We acknowledge your letter of resignation from the office of Director of ${CO(v)}, received on ____________, with effect from ${vd(v, "rdate", "date")}. The Board will take it on record at its next meeting, and the Company will file Form DIR-12 within thirty days. We thank you for your contribution.`), S(), L("Yours faithfully,"), R("For **" + CO(v) + "**"), S(), S(), R("Director / Company Secretary"));
      return o;
    },
  },
  {
    id: "l-bank", group: "Other letters", title: "Letter to the bank for opening an account", short: "A request to open a current account, with the authorised signatories and documents listed.",
    basis: "Board resolution (Section 179(3) and Companies (Meetings of Board and its Powers) Rules, 2014) and the bank's own requirements",
    fields: [{ id: "bank", label: "Name of the bank" }, { id: "branch", label: "Branch and address", type: "textarea" }, { id: "acct", label: "Type of account", ph: "Current account" }, { id: "mode", label: "Mode of operation", ph: "Jointly by any two of the authorised signatories" }, { id: "signs", label: "Authorised signatories (one per line)", type: "textarea" }, { id: "rdate", label: "Date of Board resolution", type: "date" }].concat(LF, LSG),
    notes: ["Every bank has its own account-opening form and know-your-customer papers. This letter goes with them."],
    build(v) {
      return letter(v, ["The Branch Manager,", va(v, "bank", "bank"), (v.branch || "[branch address]").replace(/\n+/g, ", ")], `Request to open a ${va(v, "acct", "current account")} in the name of ${CO(v)}`, [`The Board of Directors of the Company, at its meeting held on ${vd(v, "rdate", "date")}, resolved to open a ${va(v, "acct", "current account")} with your branch. We request you to open the account. A certified copy of the resolution is enclosed.`, "The following persons are authorised to operate the account, and their specimen signatures are on the form attached:"].concat(LN("signs")(v).map((x, i) => `${i + 1}. ${x}`)).concat([`The account will be operated: ${va(v, "mode", "mode of operation")}.`, "We enclose the following documents: certificate of incorporation; memorandum and articles of association; PAN of the Company; the Board resolution; the KYC documents and photographs of the authorised signatories; proof of the address of the registered office; and the specimen signature card."]));
    },
  },
  {
    id: "l-noc", group: "Other letters", title: "No-objection letter from the owner of the registered office", short: "The owner's consent to use the premises as the registered office (attaches to INC-22).",
    basis: "Section 12(1) and (2); Rule 25, Companies (Incorporation) Rules, 2014 (Form INC-22)",
    fields: [{ id: "owner", label: "Name of the owner of the premises" }, { id: "oaddr", label: "Address of the owner", type: "textarea" }, { id: "prem", label: "Address of the premises", type: "textarea" }, { id: "tenure", label: "Term of use", ph: "from 1 April 2026 for 11 months" }, { id: "basis", label: "Basis of use", type: "select", opts: [["on rent", "On rent"], ["free of charge", "Free of charge"], ["under a lease", "Under a lease"]] }].concat(LF),
    notes: ["File this with Form INC-22 along with a recent utility bill (not more than two months old) for the premises and the rent or lease agreement if there is one. The owner signs on their own letterhead."],
    build(v) {
      const o = header({ company: "", cin: "", ro: "" }).slice(0, 0);
      o.push(R("Date: " + vd(v, "ldate", "date")), S(), L("To,"), L("The Registrar of Companies"), S(), P("**Subject: No-objection certificate for the use of premises as the registered office of " + CO(v) + "**"), S(), L("Dear Sir,"), S());
      o.push(P(`I, ${va(v, "owner", "name")}, of ${(v.oaddr || "[address]").replace(/\n+/g, ", ")}, am the owner of the premises at ${(v.prem || "[address of premises]").replace(/\n+/g, ", ")}. I state that I have no objection to ${CO(v)} using these premises as its registered office ${va(v, "basis", "basis")}, for the term ${va(v, "tenure", "term")}.`), P("I confirm that I have given this consent of my own free will, and that the premises are in my lawful possession and can be used for this purpose."), S(), L("Yours faithfully,"), S(), S(), L("____________________"), L("**" + va(v, "owner", "Owner") + "**"), L("PAN: ____________"), L("Address: " + (v.oaddr || "[address]").replace(/\n+/g, ", ")), L("Date: ____________"));
      return o;
    },
  },
  {
    id: "l-dirdisc", group: "Letters to directors, auditors and members", title: "Request to directors for annual disclosures", short: "A letter asking each director for MBP-1 and DIR-8, with the dates.",
    basis: "Sections 184(1) and 164(2); Rule 9, Companies (Meetings of Board and its Powers) Rules, 2014; Rule 14, Companies (Appointment and Qualification of Directors) Rules, 2014",
    fields: [{ id: "fy", label: "Financial year", ph: "2026-27" }, { id: "due", label: "Please send by", type: "date" }, { id: "bmdate", label: "First Board meeting of the year", type: "date" }].concat(LF, LSG),
    notes: ["The notice in Form MBP-1 is given at the first Board meeting of every financial year and whenever it changes. Form DIR-8 is given at the first Board meeting after the appointment, and then each year."],
    build(v) {
      const o = header(v); o.push(L("Ref: " + ((v.ref || "").trim() || "[number]")), R("Date: " + vd(v, "ldate", "date")), S(), L("To,"), L("All the Directors,"), L(CO(v)), S(), P(`**Subject: Annual disclosures for the financial year ${va(v, "fy", "year")}**`), S(), L("Dear Sir / Madam,"), S());
      o.push(P(`The law needs each director to give the Company certain notices and declarations at the first Board meeting of every financial year, which will be held on ${vd(v, "bmdate", "date")}. Please send them to us by **${vd(v, "due", "date")}**:`));
      ["**Form MBP-1:** notice of disclosure of interest under Section 184(1), listing the companies, bodies corporate, firms and other associations in which you are interested, and any change since the last notice.", "**Form DIR-8:** your declaration that you are not disqualified from being a director under Section 164(2).", "**Declaration of independence** (independent directors only) under Section 149(7), with the details in Schedule IV.", "A list of the directorships and committee memberships you hold, for the Board's report and the corporate governance report (listed companies)."].forEach((x, i) => o.push(I(`${i + 1}. ${x}`)));
      o.push(P("The Company will keep the forms in the register of directors' interests and place them before the Board. Please tell us at once if anything changes during the year."), S(), L("Yours faithfully,"), R("For **" + CO(v) + "**"), S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")));
      return o;
    },
  },
  {
    id: "l-kyc", group: "Letters to directors, auditors and members", title: "Letter to members to update their details", short: "A request to update PAN, e-mail, bank details and nomination, and to dematerialise.",
    basis: "Section 88; Section 72 (nomination); Rules 9 and 10 of the Companies (Prospectus and Allotment of Securities) Rules, 2014; SEBI circulars for listed companies",
    fields: [{ id: "due", label: "Please reply by", type: "date" }, { id: "contact", label: "Whom to contact" }, { id: "email", label: "E-mail for replies" }, { id: "listed", label: "Company type", type: "select", opts: [["unlisted", "Unlisted"], ["listed", "Listed"]] }].concat(LF, LSG),
    notes: ["For an unlisted company, the Rules require every holder of securities to dematerialise before they transfer or are offered securities (check Rule 9B). For a listed company, follow the SEBI circular on KYC for physical holders."],
    build(v) {
      const o = header(v); o.push(L("Ref: " + ((v.ref || "").trim() || "[number]")), R("Date: " + vd(v, "ldate", "date")), S(), L("To,"), L("The Member"), L("[Name and address of the member]"), L("Folio no.: ____________"), S(), P("**Subject: Please update your details with the Company**"), S(), L("Dear Member,"), S());
      o.push(P(`We keep the register of members under Section 88 of the Companies Act, 2013 and need correct details for every member, so that notices, dividends and other benefits reach you. Please send us the following by **${vd(v, "due", "date")}** to ${va(v, "contact", "contact")} at ${va(v, "email", "e-mail")}:`));
      ["PAN, with a self-attested copy of the PAN card.", "E-mail address and mobile number, for sending notices and annual reports electronically.", "Bank account details (account number, name of the bank and branch, and IFSC), with a cancelled cheque, for paying dividends directly into your account.", "Nomination under Section 72 of the Act, if you have not nominated, or any change in your nomination.", "Specimen signature, if it has changed.", v.listed === "listed" ? "If you hold shares in physical form, please follow the KYC steps that SEBI requires. We encourage you to hold your shares in dematerialised form." : "If you hold shares in physical form, we encourage you to dematerialise them, which makes transfers and holding easier."].forEach((x, i) => o.push(I(`${i + 1}. ${x}`)));
      o.push(P("If your address has changed, or if you have moved your shares to someone else, please tell us."), S(), L("Yours faithfully,"), R("For **" + CO(v) + "**"), S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")));
      return o;
    },
  },
  {
    id: "l-allotadv", group: "Letters to directors, auditors and members", title: "Allotment advice to the allottee", short: "A letter telling an applicant that shares were allotted, with the details.",
    basis: "Section 39 (public issue) and Section 42 and Rule 14 (private placement); Section 56",
    fields: [{ id: "allottee", label: "Name of the allottee" }, { id: "aaddr", label: "Address", type: "textarea" }, { id: "adate", label: "Date of allotment", type: "date" }, { id: "shares", label: "Number of shares allotted", type: "number" }, { id: "class", label: "Class of shares", def: "equity shares" }, { id: "fv", label: "Face value (Rs.)", type: "number" }, { id: "price", label: "Issue price (Rs.)", type: "number" }, { id: "recd", label: "Amount received (Rs.)", type: "number" }, { id: "folio", label: "Folio number" }, { id: "dp", label: "Credit to demat account (DP ID and Client ID), if any" }].concat(LF, LSG),
    notes: ["Share certificates are due within two months of allotment (Section 56(4)) and the return of allotment in Form PAS-3 within the time the Rules allow (15 days for a private placement). Check the current rules for dematerialisation."],
    build(v) {
      return letter(v, [va(v, "allottee", "allottee"), (v.aaddr || "[address]").replace(/\n+/g, ", ")], `Allotment of ${va(v, "class", "shares")} of ${CO(v)}`, [`We are happy to inform you that the Board of Directors of the Company, at its meeting held on ${vd(v, "adate", "date")}, allotted to you **${vnum(v, "shares", "number")} ${va(v, "class", "class")}** of face value Rs. ${v.fv ? inrFmt(v.fv) : "[value]"} each, at an issue price of Rs. ${v.price ? inrFmt(v.price) : "[price]"} per share, against your application and the amount of ${v.recd ? vmoney(v, "recd") : "[amount]"} received.`, `Your folio number is ${va(v, "folio", "folio")}.${v.dp ? " The shares are being credited to your demat account " + v.dp + "." : " The share certificate will be sent to you within two months of the allotment."}`, "Please keep this letter for your records."]);
    },
  },
  {
    id: "l-unclaimed", group: "Letters to directors, auditors and members", title: "Reminder to members about unclaimed dividend", short: "A reminder before unclaimed dividend and shares go to the IEPF.",
    basis: "Sections 124 and 125, Companies Act, 2013; Investor Education and Protection Fund Authority (Accounting, Audit, Transfer and Refund) Rules, 2016",
    fields: [{ id: "fy", label: "Financial year of the dividend", ph: "2019-20" }, { id: "due", label: "Claim before", type: "date" }, { id: "iepf", label: "Date on which the amount and the shares will be transferred", type: "date" }, { id: "contact", label: "Whom to contact" }, { id: "email", label: "E-mail for claims" }].concat(LF, LSG),
    notes: ["Dividend that stays unpaid or unclaimed for seven years goes to the IEPF, and so do the shares on which dividend has not been claimed for seven consecutive years. Send individual notices to the members concerned and publish a newspaper notice, as the Rules require."],
    build(v) {
      const o = header(v); o.push(L("Ref: " + ((v.ref || "").trim() || "[number]")), R("Date: " + vd(v, "ldate", "date")), S(), L("To,"), L("The Member"), L("[Name and address]"), L("Folio no.: ____________"), S(), P("**Subject: Unclaimed dividend for the financial year " + va(v, "fy", "year") + "**"), S(), L("Dear Member,"), S());
      o.push(P(`Our records show that you have not claimed the dividend for the financial year ${va(v, "fy", "year")}. Under Section 124 of the Companies Act, 2013, dividend that remains unclaimed for seven years must be transferred to the Investor Education and Protection Fund (IEPF), and under Section 124(6) the shares on which dividend has not been paid or claimed for seven consecutive years must also be transferred to the IEPF.`), P(`Please claim your dividend before **${vd(v, "due", "date")}** by writing to ${va(v, "contact", "contact")} at ${va(v, "email", "e-mail")}, with your folio number, a copy of your PAN, a cancelled cheque and a self-attested identity proof. If we do not hear from you, we will transfer the amount and the shares to the IEPF on or after ${vd(v, "iepf", "date")}. You may still claim them from the IEPF Authority afterwards by filing Form IEPF-5.`), S(), L("Yours faithfully,"), R("For **" + CO(v) + "**"), S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")));
      return o;
    },
  },
  {
    id: "l-transferdef", group: "Letters to directors, auditors and members", title: "Deficiency memo on a share transfer", short: "A letter returning transfer papers with the defects listed.",
    basis: "Section 56, Companies Act, 2013; Rule 7, Companies (Share Capital and Debentures) Rules, 2014; Form SH-4",
    fields: [{ id: "to", label: "Name of the person who lodged the transfer" }, { id: "taddr", label: "Address", type: "textarea" }, { id: "ldt", label: "Date the transfer was lodged", type: "date" }, { id: "from", label: "Transferor" }, { id: "tt", label: "Transferee" }, { id: "shares", label: "Number of shares", type: "number" }, { id: "defects", label: "Defects (one per line)", type: "textarea", def: "Signature of the transferor does not match the specimen\nStamp duty is not paid in full" }].concat(LF, LSG),
    notes: ["The Company must send the notice of refusal or return within thirty days of receiving the instrument of transfer (Section 56(4)). The Board may refuse a transfer only for the reasons in the articles and the law."],
    build(v) {
      return letter(v, [va(v, "to", "name"), (v.taddr || "[address]").replace(/\n+/g, ", ")], `Share transfer papers returned: ${vnum(v, "shares", "number")} shares from ${va(v, "from", "transferor")} to ${va(v, "tt", "transferee")}`, [`We received the transfer deed (Form SH-4) on ${vd(v, "ldt", "date")} and the share certificate(s). We cannot register the transfer now because of the following defects:`].concat(LN("defects")(v).map((x, i) => `${i + 1}. ${x}`)).concat(["We return the papers with this letter. Please correct the defects and send the papers again. The Company will register the transfer once it has the complete papers, as the law and the articles of the Company allow."]));
    },
  },
  {
    id: "l-shortnotice", group: "Letters to directors, auditors and members", title: "Member's consent to shorter notice of a general meeting", short: "A consent from the member, with the 95% test, for a meeting on shorter notice.",
    basis: "Section 101(1) proviso, Companies Act, 2013; Secretarial Standard 2",
    fields: [{ id: "kind", label: "Meeting", type: "select", opts: [["Annual General Meeting", "Annual General Meeting"], ["Extraordinary General Meeting", "Extraordinary General Meeting"]] }, { id: "mdate", label: "Date of meeting", type: "date" }, { id: "mtime", label: "Time", ph: "11:00 a.m." }, { id: "venue", label: "Venue", type: "textarea" }, { id: "ndate", label: "Date of notice", type: "date" }, { id: "days", label: "Clear days of notice given", type: "number" }],
    notes: ["A general meeting may be called on shorter notice if members holding at least ninety-five per cent of the paid-up share capital with voting rights (or, for a company without share capital, ninety-five per cent of the total voting power) consent in writing or electronically. For an AGM, ninety-five per cent of the members entitled to vote must consent. Keep every consent with the notice, and mention the consent in the minutes."],
    build(v) {
      const o = header(v); o.push(T("CONSENT TO SHORTER NOTICE OF A GENERAL MEETING"), R("Date: ____________"), S());
      o.push(P(`I/We, the undersigned member(s) of **${CO(v)}**, entitled to vote at the ${va(v, "kind", "meeting")}, have seen the notice dated ${vd(v, "ndate", "date")}, which was given to me/us ${v.days ? inrFmt(v.days) : "[number]"} clear days before the meeting to be held on ${vd(v, "mdate", "date")} at ${va(v, "mtime", "time")} at ${va(v, "venue", "venue")}.`), P("Under the proviso to Section 101(1) of the Companies Act, 2013, I/we give my/our consent to the meeting being held on shorter notice than that which the Act requires, and I/we waive any objection to the shorter notice."));
      o.push(TB([["Name of the member", "Folio no. / DP ID and Client ID", "Number of shares held", "Signature"], ["", "", "", ""], ["", "", "", ""], ["", "", "", ""], ["", "", "", ""]], [2800, 2600, 1800, 1800]));
      return o;
    },
  },
  {
    id: "l-circular", group: "Meetings", title: "Resolution by circulation to the directors", short: "A draft resolution sent to the directors for approval, with a reply slip.",
    basis: "Section 175, Companies Act, 2013; Secretarial Standard 1 (resolutions by circulation)",
    fields: [{ id: "cdate", label: "Date of circulation", type: "date" }, { id: "last", label: "Reply by", type: "date" }, { id: "subject", label: "Subject of the resolution" }, { id: "text", label: "Text of the resolution", type: "textarea", ph: "RESOLVED THAT ..." }, { id: "why", label: "Why it cannot wait for a meeting", type: "textarea" }].concat(SG),
    notes: ["Section 175 lets the Board pass a resolution by circulation if it has gone to all the directors, and a majority of the directors entitled to vote approve it. If one third or more of the directors ask for a meeting, the matter goes to a meeting. The resolution is noted at the next Board meeting.", "The matters in Section 179(3) and Rule 8 (such as issuing securities, borrowing and approving the financial statements) cannot be passed by circulation."],
    build(v) {
      const o = header(v);
      o.push(T("RESOLUTION BY CIRCULATION"), C("[Section 175 of the Companies Act, 2013]"), R("Date: " + vd(v, "cdate", "date")), L("To all the directors"), S(), P(`**Subject: ${va(v, "subject", "subject")}**`));
      o.push(P(`The Board is asked to approve the following resolution by circulation. ${va(v, "why", "[Why it cannot wait for the next meeting]")}`), P((v.text || "**RESOLVED THAT** [text]").replace(/^RESOLVED THAT/i, "**RESOLVED THAT**")));
      o.push(P(`Please record your assent or dissent on the reply slip below and return it to the Company Secretary by **${vd(v, "last", "date")}**. If a director who is entitled to vote does not reply by then, the Company will take it that the director has not approved. If one third or more of the directors ask that the matter be decided at a meeting, it will be.`), ...SGN(v).slice(0, 4));
      o.push(S(), C("**REPLY SLIP**"), L("To: The Company Secretary, " + CO(v)), L("I have read the resolution on: " + va(v, "subject", "subject")), L("I  assent / dissent from  the resolution (strike out one)."), S(), L("Name of the director: ____________________"), L("DIN: ____________"), L("Signature: ____________________"), L("Date: ____________"));
      return o;
    },
  },
  {
    id: "l-greenmail", group: "Letters to directors, auditors and members", title: "Request for e-mail address for electronic notices", short: "A request to members to register an e-mail address, for sending notices and the annual report.",
    basis: "Section 20; Rule 35, Companies (Incorporation) Rules, 2014; Rule 18, Companies (Management and Administration) Rules, 2014",
    fields: [{ id: "email", label: "E-mail for replies" }, { id: "due", label: "Please reply by", type: "date" }].concat(LF, LSG),
    notes: ["A member may ask for a document in a particular mode, and the Company must send it in that mode if the member pays the fee set in the articles. Keep every reply."],
    build(v) {
      const o = header(v); o.push(L("Ref: " + ((v.ref || "").trim() || "[number]")), R("Date: " + vd(v, "ldate", "date")), S(), L("To,"), L("The Member"), L("[Name and address]"), S(), P("**Subject: Receive notices and the annual report by e-mail**"), S(), L("Dear Member,"), S());
      o.push(P(`The law allows us to send notices of general meetings, annual reports and other documents to you by e-mail. It saves paper and reaches you sooner. If you wish to get them this way, please send your e-mail address to ${va(v, "email", "e-mail")} by **${vd(v, "due", "date")}**, quoting your folio number or DP ID and Client ID. You may change your mind at any time by writing to us, and you may ask for a paper copy of any document free of charge.`), S(), L("Yours faithfully,"), R("For **" + CO(v) + "**"), S(), S(), R("**" + va(v, "sname", "name") + "**"), R(va(v, "sdes", "designation")));
      return o;
    },
  },
  {
    id: "l-complaint", group: "Letters to directors, auditors and members", title: "Acknowledgement of an investor complaint", short: "A reply that acknowledges an investor complaint and says how and when it will be resolved.",
    basis: "Regulation 13 and Regulation 20, SEBI (LODR) Regulations, 2015; SEBI circulars on investor grievances (SCORES)",
    fields: [{ id: "to", label: "Name of the complainant" }, { id: "taddr", label: "Address", type: "textarea" }, { id: "cno", label: "Complaint number" }, { id: "cdt", label: "Date received", type: "date" }, { id: "about", label: "What the complaint is about", type: "textarea" }, { id: "days", label: "Days within which you will resolve it", type: "number", def: "21" }, { id: "contact", label: "Contact person" }].concat(LF, LSG),
    notes: ["A listed company reports the complaints received and resolved every quarter. SEBI sets time limits for resolution, so check the current circular for your type of complaint."],
    build(v) {
      return letter(v, [va(v, "to", "name"), (v.taddr || "[address]").replace(/\n+/g, ", ")], `Your complaint no. ${va(v, "cno", "number")} dated ${vd(v, "cdt", "date")}`, [`We thank you for writing to us. We acknowledge your complaint received on ${vd(v, "cdt", "date")}, about the following: ${va(v, "about", "subject of the complaint")}`, `We have referred your complaint to the officer concerned and will send you our reply within ${v.days ? inrFmt(v.days) : "[number]"} days of receiving it. If you need to talk to us, please contact ${va(v, "contact", "contact person")} and quote the complaint number above.`, "If you are not satisfied with our reply, you may take up the matter on the SCORES portal of SEBI, or approach the appropriate forum."]);
    },
  },
);


/* ---------- blocks -> html (preview) and docx ---------- */
function inlineHtml(t) { return esc(t).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/\n/g, "<br>"); }
function blocksHtml(bs) {
  return bs.map(b => {
    if (b.k === "S") return '<p class="sp"></p>'; if (b.k === "PB") return '<hr class="pb">';
    if (b.k === "TB") return '<table class="dt">' + b.rows.map((r, i) => '<tr>' + r.map(c => (i ? '<td>' : '<th>') + inlineHtml(c) + (i ? '</td>' : '</th>')).join('') + '</tr>').join('') + '</table>';
    const c = { T: "t", H: "h", P: "p", L: "l", C: "c", R: "r", I: "i" }[b.k]; return `<p class="${c}"${b.k === "I" ? ` style="margin-left:${2 * (b.n || 1)}rem"` : ""}>${inlineHtml(b.t)}</p>`;
  }).join("");
}
function blocksText(bs) { return bs.map(b => b.k === "S" || b.k === "PB" ? "" : b.k === "TB" ? b.rows.map(r => r.join(" | ")).join("\n") : String(b.t).replace(/\*\*/g, "")).join("\n").replace(/\n{3,}/g, "\n\n"); }

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
    if (b.k === "TB") { const n = b.rows[0].length, tw = 9000, ws = b.w || b.rows[0].map(() => Math.floor(tw / n)); const bd = ["top", "left", "bottom", "right", "insideH", "insideV"].map(x => `<w:${x} w:val="single" w:sz="4" w:space="0" w:color="444444"/>`).join("");
      return `<w:tbl><w:tblPr><w:tblW w:w="${ws.reduce((a, c) => a + c, 0)}" w:type="dxa"/><w:tblBorders>${bd}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${ws.map(w => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>` +
        b.rows.map((r, i) => `<w:tr>${r.map((c, j) => `<w:tc><w:tcPr><w:tcW w:w="${ws[j]}" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="40"/><w:jc w:val="left"/></w:pPr>${runs(c, i ? { sz: 20 } : { b: 1, sz: 20 })}</w:p></w:tc>`).join("")}</w:tr>`).join("") + `</w:tbl><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>`; }
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
    nav = [("../", "Home", ""), ("../#/act", "Companies Act", ""), ("../#/lodr", "SEBI LODR", ""), ("../#/cases", "Case laws", ""), ("../#/forms", "Forms", ""), ("../#/processes", "Processes", ""),
           ("../#/resolutions", "Model resolutions", ""), ("../drafts/", "Drafts", ""), ("../generators/", "Document generators", "gen"), ("../excel/", "Excel utilities", "xl"),
           ("../calendar/", "Compliance calendar", "cal"), ("../stamp/", "Stamp duty", "stamp"), ("../#/tools", "Calculator", ""), ("../#/me", "My workspace", "")]
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


CAL_JS = r"""
"use strict";
const $ = (s, r) => (r || document).querySelector(s);
const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const store = { get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }, set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} } };
const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const D = (y, m, d) => new Date(Date.UTC(y, m, d));
const parse = s => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ""); return m ? D(+m[1], +m[2] - 1, +m[3]) : null; };
const addD = (d, n) => new Date(d.getTime() + n * 864e5);
const addM = (d, n) => { const y = d.getUTCFullYear(), m = d.getUTCMonth() + n, day = d.getUTCDate(); const t = D(y, m, 1); const last = D(t.getUTCFullYear(), t.getUTCMonth() + 1, 0).getUTCDate(); return D(t.getUTCFullYear(), t.getUTCMonth(), Math.min(day, last)); };
const addWD = (d, n) => { let x = d; while (n > 0) { x = addD(x, 1); const w = x.getUTCDay(); if (w !== 0 && w !== 6) n--; } return x; };
const fmt = d => `${d.getUTCDate()} ${MON[d.getUTCMonth()]} ${d.getUTCFullYear()}`;
const iso = d => d.toISOString().slice(0, 10);
const ymd = d => iso(d).replace(/-/g, "");
const today = (() => { const n = new Date(); return D(n.getFullYear(), n.getMonth(), n.getDate()); })();

/* Each rule: who it applies to, the due date, and where it comes from. cat: Annual, Quarterly, Half-yearly, Event. */
function annual(fyEnd, agm, c) {
  const fyStart = addD(addM(fyEnd, -12), 1), L = c.listed, out = [];
  const add = (cat, what, form, date, basis, note) => out.push({ cat, what, form, date, basis, note: note || "" });
  const mar31 = (() => { let y = fyStart.getUTCFullYear(); let d = D(y, 2, 31); if (d < fyStart) d = D(y + 1, 2, 31); return d; })();
  add("Annual", "Hold the Board meeting to approve the financial statements and the Board's report", "Board meeting", addD(agm, -22), "Sections 134 and 173", "Before the AGM notice goes out");
  add("Annual", "Send the notice of the AGM to the members", "Notice", addD(agm, -22), "Section 101", "At least 21 clear days before the meeting");
  add("Annual", "Hold the Annual General Meeting", "AGM", agm, "Section 96", "Within six months of the year end, and not more than 15 months after the last AGM");
  add("Annual", "File the financial statements", "AOC-4", addD(agm, 30), "Section 137", "Within 30 days of the AGM");
  add("Annual", "File the annual return", c.small ? "MGT-7A" : "MGT-7", addD(agm, 60), "Section 92", "Within 60 days of the AGM");
  add("Annual", "File the notice of the auditor's appointment", "ADT-1", addD(agm, 15), "Section 139; Rule 4", "Within 15 days of the AGM, if the auditor was appointed or reappointed");
  add("Annual", "File special resolutions passed at the AGM", "MGT-14", addD(agm, 30), "Section 117", "Within 30 days of the resolution");
  add("Annual", "File the return of deposits and of money not treated as deposits", "DPT-3", D(mar31.getUTCFullYear(), 5, 30), "Rule 16 and 16A", "On or before 30 June, with data as on 31 March");
  if (c.mr3) add("Annual", "Obtain the secretarial audit report", "MR-3", addD(agm, -22), "Section 204", "To go with the Board's report");
  add("Half-yearly", "Return of payments to micro and small enterprises outstanding over 45 days: April to September", "MSME-1", D(mar31.getUTCFullYear() - 1, 9, 31), "Section 405", "By 31 October");
  add("Half-yearly", "Return of payments to micro and small enterprises outstanding over 45 days: October to March", "MSME-1", D(mar31.getUTCFullYear(), 3, 30), "Section 405", "By 30 April");
  if (L) {
    for (let q = 1; q <= 4; q++) {
      const qe = addM(fyEnd, -12 + 3 * q), qn = "Q" + q, yearEnd = q === 4;
      add("Quarterly", `${qn}: financial results (quarter ended ${fmt(qe)})`, "Regulation 33", addD(qe, yearEnd ? 60 : 45), "Regulation 33, LODR", yearEnd ? "Within 60 days of the year end for the annual results" : "Within 45 days of the quarter end");
      add("Quarterly", `${qn}: shareholding pattern`, "Regulation 31", addD(qe, 21), "Regulation 31, LODR", "Within 21 days of the quarter end");
      add("Quarterly", `${qn}: corporate governance report`, "Regulation 27(2)", addD(qe, 21), "Regulation 27, LODR", "Within 21 days of the quarter end");
      add("Quarterly", `${qn}: statement of investor complaints`, "Regulation 13(3)", addD(qe, 21), "Regulation 13, LODR", "Within 21 days of the quarter end");
    }
    add("Annual", "Annual secretarial compliance report to the exchanges", "Regulation 24A", addD(fyEnd, 60), "Regulation 24A, LODR", "Within 60 days of the year end");
    add("Annual", "Send the annual report to the exchanges and the members", "Regulation 34", addD(agm, -22), "Regulation 34, LODR", "Not later than the day dispatch starts, and with the AGM notice");
    add("Annual", "Report the voting results of the AGM to the exchanges", "Regulation 44(3)", addWD(agm, 2), "Regulation 44, LODR", "Within 2 working days of the meeting. Holidays are not counted here");
  }
  return out;
}
const EVENTS = [
  ["Appointment of a director or key managerial personnel", "DIR-12", 30, "Section 170(2)", "Within 30 days"],
  ["Resignation of a director (the company's filing)", "DIR-12", 30, "Section 168", "Within 30 days of receiving the notice"],
  ["Passing of a special resolution", "MGT-14", 30, "Section 117", "Within 30 days"],
  ["Allotment of shares by private placement", "PAS-3", 15, "Rule 14, Prospectus and Allotment Rules", "Within 15 days"],
  ["Allotment of shares (other than private placement)", "PAS-3", 30, "Section 39(4)", "Within 30 days"],
  ["Issue of share certificates after an allotment", "SH-1", "M2", "Section 56(4)", "Within 2 months of the allotment"],
  ["Creation of a charge", "CHG-1", 30, "Section 77", "Within 30 days. A further 60 days with additional fee"],
  ["Satisfaction of a charge", "CHG-4", 30, "Section 82", "Within 30 days"],
  ["Change of the registered office within the same city", "INC-22", 30, "Section 12", "Within 30 days"],
  ["Incorporation: first Board meeting", "Board meeting", 30, "Section 173(2)", "Within 30 days of incorporation"],
  ["Incorporation: first auditor by the Board", "ADT-1", 30, "Section 139(6)", "Within 30 days of incorporation (members have 90 days if the Board does not)"],
  ["Incorporation: declaration of commencement of business", "INC-20A", 180, "Section 10A", "Within 180 days of incorporation"],
  ["Incorporation: issue of share certificates to the subscribers", "SH-1", "M2", "Section 56(4)", "Within 2 months of incorporation"],
  ["Auditor's resignation (the auditor's filing)", "ADT-3", 30, "Section 140(2)", "Within 30 days of the resignation"],
  ["Receipt of a declaration of significant beneficial ownership", "BEN-2", 30, "Section 90; SBO Rules", "Within 30 days of receiving BEN-1"],
  ["Board meeting held (send the draft minutes)", "Draft minutes", 15, "Secretarial Standard 1", "Within 15 days of the meeting"],
  ["Board meeting held (the chairman signs the minutes)", "Signed minutes", 30, "Secretarial Standard 1", "Within 30 days of the meeting"],
  ["Share transfer instrument lodged", "Registration or refusal", 30, "Section 56(4)(c)", "Within 30 days. Send the notice of refusal in that time"],
];
let ROWS = [], PICK = {};
function build() {
  const fyEnd = parse($("#fy").value), agm = parse($("#agm").value);
  if (!fyEnd || !agm) { $("#out").innerHTML = '<p class="small">Enter the year end and the AGM date.</p>'; return; }
  const c = { listed: $("#kind").value === "listed", small: $("#kind").value === "small", mr3: $("#mr3").checked };
  ROWS = annual(fyEnd, agm, c).sort((a, b) => a.date - b.date);
  const ev = parse($("#evd").value), ei = +$("#evk").value;
  if (ev && EVENTS[ei]) { const e = EVENTS[ei]; ROWS.push({ cat: "Event", what: e[0], form: e[1], date: e[2] === "M2" ? addM(ev, 2) : addD(ev, e[2]), basis: e[3], note: e[4] }); ROWS.sort((a, b) => a.date - b.date); }
  ROWS.forEach((r, i) => { r.id = i; if (!(i in PICK)) PICK[i] = true; });
  const cats = ["Annual", "Half-yearly", "Quarterly", "Event"];
  $("#out").innerHTML = `<p class="small">${ROWS.length} dates. Tick the ones to put in your calendar file.</p><div class="tw"><table class="tbl"><thead><tr><th></th><th>Due date</th><th>What</th><th>Form or rule</th><th>Time left</th><th>Basis and note</th></tr></thead><tbody>` +
    ROWS.map(r => { const n = Math.round((r.date - today) / 864e5); const t = n < 0 ? `<span class="over">${-n} days ago</span>` : n === 0 ? "Today" : n <= 15 ? `<span class="soon">${n} days</span>` : `${n} days`;
      return `<tr><td><input type="checkbox" data-i="${r.id}" ${PICK[r.id] ? "checked" : ""} aria-label="Include"></td><td>${fmt(r.date)}</td><td>${esc(r.what)}</td><td>${esc(r.form)}</td><td>${t}</td><td>${esc(r.basis)}${r.note ? ". " + esc(r.note) : ""}</td></tr>`; }).join("") + "</tbody></table></div>";
  store.set("cal_company", $("#co").value);
}
function ics() {
  const co = ($("#co").value || "Company").replace(/[,;\\]/g, " ");
  const L = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Company Law Knowledge Lab//Compliance calendar//EN", "CALSCALE:GREGORIAN"];
  ROWS.filter(r => PICK[r.id]).forEach((r, i) => {
    const t = s => String(s).replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
    L.push("BEGIN:VEVENT", `UID:${ymd(r.date)}-${i}-${r.id}@companylawlab`, `DTSTAMP:${ymd(today)}T000000Z`, `DTSTART;VALUE=DATE:${ymd(r.date)}`, `DTEND;VALUE=DATE:${ymd(addD(r.date, 1))}`,
      `SUMMARY:${t(co + ": " + r.form + " due: " + r.what)}`, `DESCRIPTION:${t(r.basis + ". " + r.note + ". Check the current law before you rely on this date.")}`,
      "BEGIN:VALARM", "ACTION:DISPLAY", `DESCRIPTION:${t(r.form + " is due in 7 days")}`, "TRIGGER:-P7D", "END:VALARM", "END:VEVENT");
  });
  L.push("END:VCALENDAR");
  const blob = new Blob([L.join("\r\n")], { type: "text/calendar" }); const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "compliance-calendar.ics"; document.body.appendChild(a); a.click(); a.remove();
}
function csv() {
  const q = s => '"' + String(s).replace(/"/g, '""') + '"';
  const lines = [["Due date", "What", "Form or rule", "Basis", "Note"].map(q).join(",")].concat(ROWS.filter(r => PICK[r.id]).map(r => [iso(r.date), r.what, r.form, r.basis, r.note].map(q).join(",")));
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob(["\ufeff" + lines.join("\r\n")], { type: "text/csv" })); a.download = "compliance-calendar.csv"; document.body.appendChild(a); a.click(); a.remove();
}
function init() {
  const sv = store.get("theme") || "auto"; document.documentElement.dataset.theme = sv;
  document.querySelectorAll(".sw").forEach(b => { b.setAttribute("aria-pressed", String(b.dataset.t === sv)); b.onclick = () => { document.documentElement.dataset.theme = b.dataset.t; store.set("theme", b.dataset.t); document.querySelectorAll(".sw").forEach(x => x.setAttribute("aria-pressed", String(x === b))); }; });
  const y = today.getUTCMonth() >= 3 ? today.getUTCFullYear() + 1 : today.getUTCFullYear();
  $("#fy").value = `${y}-03-31`; $("#agm").value = `${y}-09-30`; $("#co").value = store.get("cal_company") || "";
  $("#evk").innerHTML = EVENTS.map((e, i) => `<option value="${i}">${esc(e[0])}</option>`).join("");
  $("#fy").onchange = () => { const f = parse($("#fy").value); if (f) $("#agm").value = iso(addM(f, 6)); build(); };
  ["agm", "kind", "mr3", "evd", "evk", "co"].forEach(i => { $("#" + i).oninput = build; $("#" + i).onchange = build; });
  $("#out").addEventListener("change", e => { const i = e.target.dataset && e.target.dataset.i; if (i != null) PICK[i] = e.target.checked; });
  $("#dl").onclick = ics; $("#dc").onclick = csv; $("#pr").onclick = () => window.print();
  build();
}
init();
"""
CAL_CSS = r"""
.cols{grid-template-columns:minmax(17rem,22rem) minmax(0,1fr)}
.tw{overflow-x:auto}
.over{color:#a32514;font-weight:600}.soon{color:#8a5a00;font-weight:600}
:root[data-theme=dark] .over{color:#ff9a8a}:root[data-theme=dark] .soon{color:#f0c060}
.tbl td:first-child,.tbl th:first-child{width:2rem;text-align:center}
@media print{.form,.actions{display:none!important}}
"""


def build_calendar():
    out = SITE / "calendar"
    out.mkdir(parents=True, exist_ok=True)
    body = """<h1>Compliance calendar</h1>
<p class="lede">Enter your year end and AGM date. The page lists the company-law filings and meetings with their due dates, and gives you a file for Google Calendar, Outlook or your phone, with a reminder seven days before each date.</p>
<div class="cols"><form class="form" onsubmit="return false" autocomplete="off"><h2>Your company</h2>
<div class="fld"><label for="co">Company name (for the calendar file)</label><input id="co" type="text" placeholder="ABC Private Limited"></div>
<div class="fld"><label for="kind">Type of company</label><select id="kind"><option value="private">Private or unlisted public company</option><option value="small">Small company or One Person Company</option><option value="listed">Listed company</option></select></div>
<div class="fld"><label><input id="mr3" type="checkbox" style="width:auto"> A secretarial audit applies (Section 204)</label></div>
<div class="fld"><label for="fy">Financial year ends on</label><input id="fy" type="date"></div>
<div class="fld"><label for="agm">AGM date (planned or held)</label><input id="agm" type="date"></div>
<h2>One event</h2><p class="small">Add the deadline that follows from one event, such as an allotment or an appointment.</p>
<div class="fld"><label for="evk">Event</label><select id="evk"></select></div>
<div class="fld"><label for="evd">Date of the event</label><input id="evd" type="date"></div></form>
<div><div class="actions"><button class="btn pri" id="dl" type="button">Download calendar file (.ics)</button><button class="btn" id="dc" type="button">Download as spreadsheet (.csv)</button><button class="btn" id="pr" type="button">Print</button></div>
<div id="out"></div>
<div class="check"><h2>Please note</h2><p class="small">The dates are the dates the Act, the Rules and the SEBI Regulations give, counted in calendar days from the event, and they assume a financial year ending on 31 March for the MSME-1 and DPT-3 dates. MCA and SEBI often extend dates or change them by circular. Public holidays are not counted. Check every date against the current law before you rely on it. The page does not store what you enter, except the company name in your browser. It is not legal advice.</p></div></div></div>"""
    (out / "index.html").write_text(shell("Compliance calendar | Company Law Knowledge Lab", "Company-law due dates for your year end and AGM date, with a calendar file you can download.", "cal", body, CAL_JS).replace("</style>", CAL_CSS + "</style>", 1), encoding="utf-8")
    print("Compliance calendar page written")


STAMP_JS = r"""
"use strict";
const $ = (s, r) => (r || document).querySelector(s);
const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const store = { get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }, set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} } };
const LK = 100000, CR = 10000000;
const inr = n => { n = String(Math.round(n)); if (n.length <= 3) return n; let r = n.slice(0, -3).replace(/\B(?=(\d{2})+(?!\d))/g, ","); return r + "," + n.slice(-3); };
/* Rule types. fix: a fixed amount. thr: lo up to a limit, hi above it. pct: rate of capital, with a minimum and a maximum.
   blk: an amount for each block of capital or part of a block, with an optional maximum. ker: Kerala's own scale. chk: not worked out here. */
const fix = v => ({ t: "fix", v }), thr = (l, lo, hi) => ({ t: "thr", l, lo, hi }), pct = (r, min, max) => ({ t: "pct", r, min: min || 0, max: max || 0 }), blk = (a, s, max, min) => ({ t: "blk", a, s, max: max || 0, min: min || 0 });
const NIL = { t: "nil" }, KER = { t: "ker" }, chk = x => ({ t: "chk", x });
const S = [
  // name, form fee (INC-2, INC-7, INC-29), MoA, AoA on authorised capital, SH-7 on the increase, text of the AoA rule, text of the SH-7 rule
  ["Andaman and Nicobar", 20, 200, fix(300), NIL, "Rs. 300", "Nil"],
  ["Andhra Pradesh", 20, 500, pct(.0015, 1000, 5 * LK), pct(.0015, 1000, 5 * LK), "0.15% of authorised capital, minimum Rs. 1,000, maximum Rs. 5 lakh", "0.15% of the increase, minimum Rs. 1,000, maximum Rs. 5 lakh"],
  ["Arunachal Pradesh", 10, 200, fix(500), NIL, "Rs. 500", "Nil"],
  ["Assam", 15, 200, fix(310), NIL, "Rs. 310", "Nil"],
  ["Bihar", 20, 500, pct(.0015, 1000, 5 * LK), pct(.0015, 1000, 5 * LK), "0.15% of authorised capital or Rs. 1,000, whichever is more, maximum Rs. 5 lakh", "The higher of Rs. 1,000 and the same rate on the increase, maximum Rs. 5 lakh"],
  ["Chandigarh", 3, 500, fix(1000), NIL, "Rs. 1,000", "Nil"],
  ["Chhattisgarh", 10, 500, pct(.0015, 1000, 5 * LK), pct(.0015, 1000, 5 * LK), "0.15% of authorised capital or Rs. 1,000, whichever is more, maximum Rs. 5 lakh", "The higher of Rs. 1,000 and the same rate on the increase"],
  ["Dadra and Nagar Haveli", 1, 15, fix(25), NIL, "Rs. 25", "Nil"],
  ["Daman and Diu", 20, 150, blk(1000, 5 * LK), blk(1000, 5 * LK), "Rs. 1,000 for each Rs. 5 lakh of capital or part of it", "Rs. 1,000 for each Rs. 5 lakh of the increase or part of it"],
  ["Delhi", 10, 200, pct(.0015, 0, 25 * LK), pct(.0015, 0, 25 * LK), "0.15% of authorised capital, maximum Rs. 25 lakh", "0.15% of the increase, maximum Rs. 25 lakh"],
  ["Goa", 50, 150, blk(1000, 5 * LK), blk(1000, 5 * LK), "Rs. 1,000 for each Rs. 5 lakh of capital or part of it", "Rs. 1,000 for each Rs. 5 lakh of the increase or part of it"],
  ["Gujarat", 20, 100, pct(.005, 0, 5 * LK), chk("The MCA table gives a formula based on the increase, with a maximum of Rs. 5 lakh. Read the Gujarat rule."), "0.5% of authorised capital, maximum Rs. 5 lakh", "A formula based on the increase, maximum Rs. 5 lakh"],
  ["Haryana", 15, 60, thr(LK, 60, 120), NIL, "Rs. 60 if capital is up to Rs. 1 lakh; Rs. 120 above that", "Nil"],
  ["Himachal Pradesh", 3, 60, thr(LK, 60, 120), NIL, "Rs. 60 if capital is up to Rs. 1 lakh; Rs. 120 above that", "Nil"],
  ["Jammu and Kashmir", 10, 150, thr(LK, 150, 300), NIL, "Rs. 150 if capital is up to Rs. 1 lakh; Rs. 300 above that", "Nil"],
  ["Jharkhand", 5, 63, fix(105), NIL, "Rs. 105", "Nil"],
  ["Karnataka", 20, 1000, blk(500, 10 * LK), blk(500, 10 * LK, 0, 500), "Rs. 500 for each Rs. 10 lakh of capital or part of it", "Rs. 500 for each Rs. 10 lakh of the increase or part of it, minimum Rs. 500"],
  ["Kerala", 25, 1000, KER, NIL, "Rs. 2,000 up to Rs. 10 lakh; Rs. 5,000 from Rs. 10 lakh to Rs. 25 lakh; 0.5% above Rs. 25 lakh", "Nil"],
  ["Lakshadweep", 25, 500, fix(1000), NIL, "Rs. 1,000", "Nil"],
  ["Madhya Pradesh", 50, 2500, pct(.0015, 5000, 25 * LK), chk("The MCA table gives a formula on the increase (minimum Rs. 5,000, maximum Rs. 25 lakh) that depends on whether the last increase was before or after 16 September 2014. Read the Madhya Pradesh rule."), "0.15% of authorised capital, minimum Rs. 5,000, maximum Rs. 25 lakh", "A formula on the increase, depending on the date of the last increase"],
  ["Maharashtra", 100, 200, blk(1000, 5 * LK, 50 * LK), { t: "mh" }, "Rs. 1,000 for each Rs. 5 lakh of capital or part of it, maximum Rs. 50 lakh", "Rs. 1,000 for each Rs. 5 lakh of the increase or part of it, maximum Rs. 50 lakh. No SH-7 duty once authorised capital exceeds Rs. 250 crore"],
  ["Manipur", 10, 100, fix(150), NIL, "Rs. 150", "Nil"],
  ["Meghalaya", 10, 100, fix(300), NIL, "Rs. 300", "Nil"],
  ["Mizoram", 10, 100, fix(150), NIL, "Rs. 150", "Nil"],
  ["Nagaland", 10, 100, fix(150), NIL, "Rs. 150", "Nil"],
  ["Odisha", 10, 300, fix(300), NIL, "Rs. 300", "Nil"],
  ["Puducherry", 10, 200, fix(300), NIL, "Rs. 300", "Nil"],
  ["Punjab", 25, 5000, thr(LK, 5000, 10000), NIL, "Rs. 5,000 if capital is up to Rs. 1 lakh; Rs. 10,000 above that", "Nil"],
  ["Rajasthan", 10, 500, pct(.005, 0, 0), pct(.002, 0, 25 * LK), "0.5% of authorised capital", "0.2% of the increase, maximum Rs. 25 lakh"],
  ["Tamil Nadu", 20, 200, fix(300), NIL, "Rs. 300", "Nil"],
  ["Telangana", 20, 500, pct(.0015, 1000, 5 * LK), pct(.0015, 1000, 5 * LK), "0.15% of authorised capital, minimum Rs. 1,000, maximum Rs. 5 lakh", "0.15% of the increase, minimum Rs. 1,000, maximum Rs. 5 lakh"],
  ["Tripura", 10, 100, fix(150), NIL, "Rs. 150", "Nil"],
  ["Uttar Pradesh", 10, 500, fix(500), NIL, "Rs. 500", "Nil"],
  ["Uttarakhand", 10, 500, fix(500), NIL, "Rs. 500", "Nil"],
  ["West Bengal", 10, 60, fix(300), NIL, "Rs. 300", "Nil"],
];
function calc(rule, cap) {
  if (cap <= 0) return 0;
  switch (rule.t) {
    case "fix": return rule.v;
    case "nil": return 0;
    case "thr": return cap <= rule.l ? rule.lo : rule.hi;
    case "pct": { let v = cap * rule.r; if (rule.min) v = Math.max(v, rule.min); if (rule.max) v = Math.min(v, rule.max); return v; }
    case "blk": { let v = rule.a * Math.ceil(cap / rule.s); if (rule.min) v = Math.max(v, rule.min); if (rule.max) v = Math.min(v, rule.max); return v; }
    case "ker": return cap <= 10 * LK ? 2000 : cap <= 25 * LK ? 5000 : cap * 0.005;
    default: return null;
  }
}
function sh7(s, oldCap, newCap) {
  const rule = s[4], inc = newCap - oldCap; if (inc <= 0) return 0;
  if (rule.t === "chk") return null;
  if (rule.t === "mh") { const paid = calc(s[3], oldCap), cap = 50 * LK; return Math.max(0, Math.min(calc(s[3], inc), cap - paid)); }
  return calc(rule, inc);
}
function parseAmt(id, unit) { const v = parseFloat(String($(id).value).replace(/,/g, "")); return isNaN(v) ? 0 : v * (+$(unit).value); }
function run() {
  const mode = $("#mode").value; $("#oldrow").hidden = mode !== "inc";
  const st = S[+$("#st").value]; const cap = parseAmt("#cap", "#cu"), old = parseAmt("#old", "#ou");
  store.set("stamp_state", $("#st").value);
  let h = `<h2>${esc(st[0])}</h2>`;
  if (mode === "inc") {
    const d = sh7(st, old, cap);
    h += d === null ? `<p class="warn">${esc(st[4].x)}</p>` : `<table class="tbl"><tbody><tr><th>Authorised capital before</th><td>Rs. ${inr(old)}</td></tr><tr><th>Authorised capital after</th><td>Rs. ${inr(cap)}</td></tr><tr><th>Increase</th><td>Rs. ${inr(Math.max(0, cap - old))}</td></tr><tr><th>Stamp duty on the increase (Form SH-7)</th><td><b>Rs. ${inr(d)}</b></td></tr></tbody></table><p class="small">Rule used: ${esc(st[6])}.</p>`;
  } else {
    const m = st[2], a = calc(st[3], cap), f = st[1], tot = m + a + f;
    h += `<table class="tbl"><tbody><tr><th>Authorised capital</th><td>Rs. ${inr(cap)}</td></tr><tr><th>Stamp duty on the form (INC-2, INC-7 or INC-29)</th><td>Rs. ${inr(f)}</td></tr><tr><th>Stamp duty on the Memorandum of Association</th><td>Rs. ${inr(m)}</td></tr><tr><th>Stamp duty on the Articles of Association</th><td>Rs. ${inr(a)}</td></tr><tr><th>Total</th><td><b>Rs. ${inr(tot)}</b></td></tr></tbody></table><p class="small">Rule for the Articles: ${esc(st[5])}.</p>`;
  }
  $("#res").innerHTML = h;
  const rows = S.map(s => mode === "inc" ? [s[0], sh7(s, old, cap)] : [s[0], s[1] + s[2] + calc(s[3], cap)]).sort((a, b) => (a[1] === null) - (b[1] === null) || a[1] - b[1]);
  $("#cmp").innerHTML = `<h2>All states for this amount</h2><div class="tw"><table class="tbl"><thead><tr><th>State or Union Territory</th><th>${mode === "inc" ? "Duty on the increase" : "Total duty"}</th></tr></thead><tbody>` + rows.map(r => `<tr><td>${esc(r[0])}</td><td>${r[1] === null ? "Read the state rule" : "Rs. " + inr(r[1])}</td></tr>`).join("") + "</tbody></table></div>";
}
function init() {
  const sv = store.get("theme") || "auto"; document.documentElement.dataset.theme = sv;
  document.querySelectorAll(".sw").forEach(b => { b.setAttribute("aria-pressed", String(b.dataset.t === sv)); b.onclick = () => { document.documentElement.dataset.theme = b.dataset.t; store.set("theme", b.dataset.t); document.querySelectorAll(".sw").forEach(x => x.setAttribute("aria-pressed", String(x === b))); }; });
  $("#st").innerHTML = S.map((s, i) => `<option value="${i}">${esc(s[0])}</option>`).join("");
  const keep = store.get("stamp_state"); if (keep && S[+keep]) $("#st").value = keep;
  $("#rates").innerHTML = S.map(s => `<tr><td>${esc(s[0])}</td><td>Rs. ${inr(s[1])}</td><td>Rs. ${inr(s[2])}</td><td>${esc(s[5])}</td><td>${esc(s[6])}</td></tr>`).join("");
  document.querySelectorAll("#f input,#f select").forEach(e => { e.oninput = run; e.onchange = run; });
  run();
}
init();
"""
STAMP_CSS = r"""
.cols{grid-template-columns:minmax(17rem,22rem) minmax(0,1fr)}
.tw{overflow-x:auto}.warn{background:var(--page);border:1px solid var(--rule);border-left:4px solid var(--gold);padding:.7rem 1rem}
.fld .row{display:flex;gap:.5rem}.fld .row input{flex:1}.fld .row select{width:7.5rem}
.tbl th[scope],.tbl tbody th{width:55%;background:var(--page);color:var(--ink);font-weight:600}
"""


def build_stamp():
    out = SITE / "stamp"
    out.mkdir(parents=True, exist_ok=True)
    units = '<option value="1">Rs.</option><option value="100000" selected>Lakh</option><option value="10000000">Crore</option>'
    body = f"""<h1>Stamp duty on company incorporation and capital increase</h1>
<p class="lede">Pick the state, say whether you are incorporating a company or increasing its authorised capital, and enter the amount. The page works out the stamp duty on the form, the Memorandum and the Articles, or on Form SH-7.</p>
<div class="cols"><form class="form" id="f" onsubmit="return false" autocomplete="off"><h2>What you need</h2>
<div class="fld"><label for="st">State or Union Territory (of the registered office)</label><select id="st"></select></div>
<div class="fld"><label for="mode">What for</label><select id="mode"><option value="inc0">Incorporation: MoA and AoA</option><option value="inc">Increase in authorised capital: Form SH-7</option></select></div>
<div class="fld"><label for="cap">Authorised capital (after the increase, for SH-7)</label><div class="row"><input id="cap" type="text" inputmode="decimal" value="10"><select id="cu" aria-label="Unit">{units}</select></div></div>
<div class="fld" id="oldrow" hidden><label for="old">Authorised capital before the increase</label><div class="row"><input id="old" type="text" inputmode="decimal" value="1"><select id="ou" aria-label="Unit">{units}</select></div></div></form>
<div><div id="res"></div>
<div class="check"><h2>Where the rates come from</h2><p class="small">The rates are from the e-stamp rate table of the Ministry of Corporate Affairs (<a href="https://www.mca.gov.in/mca/documents/eStamp_rate.pdf">mca.gov.in, eStamp rate</a>), for companies with share capital other than Section 8 companies, which I read on 20 November 2026. The table has no date and the MCA itself asks users to verify against the state's Stamp Act, and some states have amended their Acts since (Maharashtra, for one, has amended its Act in 2024 and has a 2026 Bill). Check the current rate with the state stamp office before you pay. Duty on a company is paid on the MCA portal along with the incorporation form.</p></div>
<div class="check"><h2>LLP agreements are not included</h2><p class="small">The MCA table covers companies only. The duty on an LLP agreement is set by each state, and the private websites I compared disagree with the MCA table on several states for companies (for example Uttar Pradesh, Punjab, Kerala and Madhya Pradesh), so I have not copied their LLP tables. Give me an official state notification for the states you need and I will add them.</p></div>
<div id="cmp"></div>
<h2>The rate table</h2><div class="tw"><table class="tbl"><thead><tr><th>State or Union Territory</th><th>Form (INC-2, 7, 29)</th><th>MoA</th><th>AoA, on authorised capital</th><th>SH-7, on the increase</th></tr></thead><tbody id="rates"></tbody></table></div></div></div>"""
    html = shell("Stamp duty on companies | Company Law Knowledge Lab", "Stamp duty on the MoA, AoA and capital increase of a company, by state, from the MCA's rate table.", "stamp", body, STAMP_JS).replace("</style>", STAMP_CSS + "</style>", 1)
    html = html.replace('<option value="inc0">', '<option value="inc0">')
    (out / "index.html").write_text(html, encoding="utf-8")
    print("Stamp duty page written")


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
    SHORT = {"register-of-directors-and-kmp.xlsx": "Directors and KMP", "register-of-members.xlsx": "Members", "register-of-charges.xlsx": "Charges", "register-of-loans-guarantees-investments.xlsx": "Loans and investments",
             "register-of-related-party-contracts.xlsx": "Related party contracts", "register-of-allotments.xlsx": "Allotments", "share-transfer-register.xlsx": "Share transfers", "register-of-debenture-holders.xlsx": "Debenture holders",
             "register-of-renewed-duplicate-certificates.xlsx": "SH-2 certificates", "register-of-dividends.xlsx": "Dividends", "register-of-deposits.xlsx": "Deposits", "register-of-significant-beneficial-owners.xlsx": "Beneficial owners BEN-3",
             "register-of-sweat-equity-shares.xlsx": "Sweat equity SH-3", "register-of-employee-stock-options.xlsx": "ESOP SH-6", "register-of-buy-back.xlsx": "Buy-back SH-10", "register-of-investments-not-in-company-name.xlsx": "Investments MBP-3",
             "register-of-director-disclosures.xlsx": "Director interests MBP-1"}
    COMBO = Workbook()
    COMBO.remove(COMBO.active)
    COMBO_ITEMS = []

    def draw_register(ws, name, basis, cols, rows, validations, extra_top):
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
                fmt = DATEF if h.startswith(("Last date", "Certificate due", "Transfer to", "Expected date", "Date of repayment due", "Lock-in ends", "Earliest vesting")) else "0.00%" if h.startswith("%") else "#,##0.00" if "(Rs.)" in h else None
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

    def register(fn, name, blurb, basis, cols, readme_lines, rows=300, validations=(), extra_top=None, combo=True):
        wb = Workbook()
        ws = wb.active
        ws.title = "Register"
        draw_register(ws, name, basis, cols, rows, validations, extra_top)
        readme(wb, name, basis, readme_lines)
        save(wb, fn, name, blurb, basis, ["Register"])
        if combo:
            sn = SHORT.get(fn, name[:28])
            draw_register(COMBO.create_sheet(sn), name, basis, cols, rows, validations, extra_top)
            COMBO_ITEMS.append((name, basis, sn))
        return ws

    register("register-of-directors-and-kmp.xlsx", "Register of directors and key managerial personnel and their shareholding", "The particulars the Act requires for each director and each key managerial person, with a column that shows who is in office now.",
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

    register("register-of-loans-guarantees-investments.xlsx", "Register of loans, guarantee, security and acquisition made by the company", "Every loan, guarantee, security and investment the company makes, with the approvals and a running total.", "Section 186(9); Rule 12, Companies (Meetings of Board and its Powers) Rules, 2014 (Form MBP-2)",
             [("Date of making", 14, "date"), ("Type", 18, "text"), ("Name of the body corporate or person", 30, "text"), ("Amount (Rs.)", 16, "num"), ("Purpose", 28, "wrap"), ("Date of Board resolution", 14, "date"),
              ("Date of special resolution (if needed)", 16, "date"), ("Rate of interest (% a year)", 12, "num"), ("Date of repayment or release", 14, "date"), ("Total so far (Rs.)", 16, "f", '=IF(E{r}="","",SUM($E${first}:E{r}))'), ("Remarks", 26, "wrap")],
             ["1. Add a row for each transaction in the order it was made. Pick the type from the list.", "2. Type the dates of the Board resolution and, where needed, the special resolution of the members.", "3. The 'Total so far' column helps you test the limits in Section 186(2). Compare it with 60% of the paid-up capital, free reserves and securities premium, or 100% of free reserves and securities premium, whichever is higher."],
             validations=[(3, ["Loan", "Guarantee", "Security", "Investment"])])

    register("register-of-related-party-contracts.xlsx", "Register of contracts or arrangements in which directors are interested", "Contracts and arrangements in which a director is interested, with the approval dates.", "Section 189; Rule 16, Companies (Meetings of Board and its Powers) Rules, 2014 (Form MBP-4)",
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
    register("share-transfer-register.xlsx", "Register of transfers of shares", "Transfers and transmissions of shares, with the date by which the new certificate is due.", "Sections 56 and 58; Secretarial Standard on transfer; Rule 7 (transfer and transmission)",
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
    register("register-of-deposits.xlsx", "Register of deposits", "Deposits accepted by the company, with the interest and the dates of repayment.", "Sections 73 to 76; Rule 14, Companies (Acceptance of Deposits) Rules, 2014",
             [("Name of depositor", 28, "text"), ("Address", 32, "wrap"), ("Date of receipt", 14, "date"), ("Amount (Rs.)", 16, "num"), ("Tenure in months", 11, "num"), ("Rate of interest (% a year)", 12, "num"), ("Date of repayment due", 15, "f", '=IF(OR(D{r}="",F{r}=""),"",EDATE(D{r},F{r}))'),
              ("Amount repaid (Rs.)", 16, "num"), ("Date of repayment", 14, "date"), ("Interest paid (Rs.)", 16, "num"), ("Date of renewal", 14, "date"), ("Remarks", 24, "wrap")],
             ["1. Add a row for each deposit. A private company may take deposits from its members only, within the limits and conditions of the Rules.", "2. The repayment date is the date of receipt plus the tenure. Type the actual date of repayment when it is paid.",
              "3. File the return of deposits in Form DPT-3 by 30 June each year, and keep the register at the registered office."],
             extra_top=[("Total received (Rs.)", '=SUM(E6:E305)')])
    register("register-of-significant-beneficial-owners.xlsx", "Register of significant beneficial owners in the company", "Persons who hold or control a significant stake, with the date by which the company must file Form BEN-2.", "Section 90; Companies (Significant Beneficial Owners) Rules, 2018 (Form BEN-3)",
             [("Name of the beneficial owner", 28, "text"), ("Date of birth", 13, "date"), ("Nationality", 13, "text"), ("Address", 32, "wrap"), ("PAN or other ID", 16, "text"), ("Date the declaration (BEN-1) was received", 16, "date"),
              ("Last date to file BEN-2 (30 days)", 15, "f", '=IF(G{r}="","",G{r}+30)'), ("BEN-2 filed on", 14, "date"), ("Name of the registered owner or the company through which held", 32, "wrap"), ("Nature of the interest", 24, "text"),
              ("Percentage of shares", 11, "num"), ("Percentage of voting rights", 11, "num"), ("Date of acquiring the interest", 14, "date"), ("Date of any change", 14, "date"), ("Date of ceasing", 14, "date"), ("Remarks", 24, "wrap")],
             ["1. Add a row for each significant beneficial owner. A person is significant if he or she holds, directly or indirectly, at least 10% of the shares or voting rights, or has the right to exercise significant influence or control.",
              "2. The company must file Form BEN-2 within 30 days of receiving the declaration. The date is worked out for you.", "3. Check the current Rules for the definitions and the thresholds."],
             validations=[(11, ["Shares", "Voting rights", "Dividend rights", "Significant influence or control", "Other"])])
    register("register-of-sweat-equity-shares.xlsx", "Register of sweat equity shares", "Shares issued for know-how or value additions, with the lock-in end date.", "Section 54; Rule 8, Companies (Share Capital and Debentures) Rules, 2014 (Form SH-3)",
             [("Name of the allottee", 28, "text"), ("Address", 32, "wrap"), ("Date of allotment", 14, "date"), ("Number of shares", 12, "num"), ("Face value (Rs.)", 12, "num"), ("Issue price (Rs.)", 12, "num"), ("Consideration (know-how, intellectual property or value additions)", 30, "wrap"),
              ("Value of the consideration (Rs.)", 16, "num"), ("Date of special resolution", 14, "date"), ("Date of the valuation report", 14, "date"), ("Lock-in ends (3 years from allotment)", 15, "f", '=IF(D{r}="","",EDATE(D{r},36))'), ("Certificate no.", 12, "text"), ("Remarks", 24, "wrap")],
             ["1. Add a row for each allotment of sweat equity shares.", "2. These shares are locked in for three years from the date of allotment. The end date is worked out for you.", "3. A special resolution and a valuation by a registered valuer are needed. Check the current Rule for the other conditions."])
    register("register-of-employee-stock-options.xlsx", "Register of employee stock options", "Options granted, vested and exercised, with the earliest vesting date.", "Section 62(1)(b); Rule 12, Companies (Share Capital and Debentures) Rules, 2014 (Form SH-6)",
             [("Name of the employee or director", 28, "text"), ("Designation", 20, "text"), ("Date of grant", 14, "date"), ("Number of options granted", 12, "num"), ("Exercise price per share (Rs.)", 14, "num"), ("Earliest vesting date (1 year from grant)", 16, "f", '=IF(D{r}="","",EDATE(D{r},12))'),
              ("Actual vesting date", 14, "date"), ("Number vested", 12, "num"), ("Date of exercise", 14, "date"), ("Number exercised", 12, "num"), ("Shares allotted", 12, "num"), ("Options lapsed or cancelled", 12, "num"), ("Options outstanding", 12, "f", '=IF(E{r}="","",E{r}-N(K{r})-N(M{r}))'),
              ("Date of special resolution", 14, "date"), ("Remarks", 24, "wrap")],
             ["1. Add a row for each grant. The scheme needs a special resolution of the members.", "2. There must be a gap of at least one year between the grant and the vesting of an option. The earliest date is worked out for you.", "3. 'Options outstanding' is the number granted, less the number exercised and the number lapsed."])
    register("register-of-buy-back.xlsx", "Register of shares or other securities bought back", "Each buy-back, with the dates for extinguishing the shares and for filing the return.", "Section 68; Rule 17, Companies (Share Capital and Debentures) Rules, 2014 (Form SH-10)",
             [("Date of Board or special resolution", 15, "date"), ("Class of securities", 16, "text"), ("Number bought back", 12, "num"), ("Price per security (Rs.)", 13, "num"), ("Total amount (Rs.)", 16, "f", '=IF(OR(D{r}="",E{r}=""),"",D{r}*E{r})'),
              ("Source of funds", 22, "text"), ("Date of declaration of solvency (SH-9)", 15, "date"), ("Date the buy-back was completed", 15, "date"), ("Last date to extinguish the securities (7 days)", 16, "f", '=IF(I{r}="","",I{r}+7)'),
              ("Date they were extinguished", 14, "date"), ("Last date to file the return (SH-11, 30 days)", 16, "f", '=IF(I{r}="","",I{r}+30)'), ("Return filed on", 14, "date"), ("Remarks", 24, "wrap")],
             ["1. Add a row for each buy-back.", "2. The securities bought back must be extinguished and destroyed within seven days of the last date of completion, and a return in Form SH-11 filed within thirty days (Section 68).",
              "3. The company cannot issue the same kind of securities for six months after a buy-back, except in the cases the Act allows."])
    register("register-of-investments-not-in-company-name.xlsx", "Register of investments not held in the company's name", "Investments that are held by someone else for the company, with the reason.", "Section 187(4); Rule 14, Companies (Meetings of Board and its Powers) Rules, 2014 (Form MBP-3)",
             [("Date of the investment", 14, "date"), ("Name of the body corporate invested in", 30, "text"), ("Description of the investment", 28, "wrap"), ("Number of shares or amount (Rs.)", 16, "num"), ("Name in which it is held", 26, "text"), ("Reason", 28, "wrap"),
              ("Date of Board resolution", 14, "date"), ("Date it was transferred to the company's name", 16, "date"), ("Remarks", 24, "wrap")],
             ["1. List every investment of the company that is held in the name of another person or body.", "2. Section 187 requires the investment to be held in the company's own name, with the exceptions the section gives. Record the reason and the Board approval."])
    register("register-of-director-disclosures.xlsx", "Register of directors' disclosures of interest", "The notices of interest, declarations and consents received from each director, with the dates.", "Sections 184, 164 and 152; Forms MBP-1, DIR-8 and DIR-2",
             [("Name of the director", 28, "text"), ("DIN", 12, "text"), ("Financial year", 12, "text"), ("Date MBP-1 was received", 14, "date"), ("Entities in which the director is interested", 36, "wrap"), ("Date of the last change", 14, "date"), ("Date tabled at the Board meeting", 15, "date"),
              ("Date DIR-8 was received (each year)", 15, "date"), ("Date of consent to act (DIR-2)", 15, "date"), ("Remarks", 24, "wrap")],
             ["1. Add a row for each director in each financial year. Record the date each document was received.", "2. Every director gives MBP-1 at the first Board meeting of each financial year, and again when the interest changes (Section 184).",
              "3. DIR-8 is given by each director every year, and DIR-2 at the time of the appointment."])

    # all-in-one workbook
    idx = COMBO.create_sheet("Index", 0)
    title(idx, "Company secretarial registers: all in one file", "One sheet for each register. Fill in the company details at the top of each sheet.")
    head(idx, 4, ["Sr", "Register", "Provision and form", "Sheet"], [5, 54, 80, 28])
    for i, (nm, bs, sn) in enumerate(COMBO_ITEMS):
        r = 5 + i
        idx.cell(row=r, column=1, value=i + 1)
        idx.cell(row=r, column=2, value=nm)
        idx.cell(row=r, column=3, value=bs)
        c = idx.cell(row=r, column=4, value=sn)
        c.hyperlink = f"#'{sn}'!A1"
        c.font = Font(color="1F4FB5", underline="single")
    style_rows(idx, 5, 4 + len(COMBO_ITEMS), 4, formula_cols=())
    for r in range(5, 5 + len(COMBO_ITEMS)):
        for c in (1, 2, 3, 4):
            idx.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFFFFF")
        idx.cell(row=r, column=4).font = Font(color="1F4FB5", underline="single")
    idx.cell(row=6 + len(COMBO_ITEMS), column=2, value="Check each register against the current Act, Rules and forms before relying on it. Free starting file from the Company Law Knowledge Lab. For study only; not legal advice.").font = Font(italic=True, color="4B5A70")
    COMBO.save(out / "company-secretarial-registers-all-in-one.xlsx")
    made.append({"file": "company-secretarial-registers-all-in-one.xlsx", "name": "All statutory registers in one file", "blurb": f"{len(COMBO_ITEMS)} registers in one workbook, with an index that links to each sheet.", "basis": "Companies Act, 2013 and the Rules",
                 "sheets": [x[2] for x in COMBO_ITEMS], "group": "Statutory registers"})

    GRP[0] = "Practice"
    register("register-of-articled-assistants.xlsx", "Register of articled assistants", "Articleship records, with the ICAI forms and the dates they fall due.", "Chartered Accountants Regulations, 1988 (ICAI Forms 102, 103, 108 and 109)",
             [("Name of the articled assistant", 28, "text"), ("Registration no.", 15, "text"), ("Date of registration", 14, "date"), ("Date of commencement", 14, "date"), ("Period in months", 10, "num"), ("Expected date of completion", 15, "f", '=IF(E{r}="","",EDATE(E{r},F{r}))'),
              ("Stipend per month (Rs.)", 14, "num"), ("Form 102 filed on", 14, "date"), ("Form 103 filed on", 14, "date"), ("Form 108 filed on", 14, "date"), ("Form 109 filed on", 14, "date"), ("Leave taken (days)", 11, "num"), ("Remarks", 24, "wrap")],
             ["1. Add a row for each articled assistant. Enter the period of articleship in months.", "2. The expected date of completion is the commencement date plus the period. Add the leave beyond the allowed days yourself.", "3. Check the current ICAI regulations for the forms and the time limits."], combo=False)

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

    # ======================================================================================
    # Batch 2: filing attachment checklists, due diligence, MR-3 checklist, accounts workbooks
    # ======================================================================================
    from openpyxl.worksheet.datavalidation import DataValidation as _DV  # noqa: F401

    def red_amber_green(ws, rng, first_cell, red=("Gap found",), amber=("Not provided",), green=("Verified", "Yes")):
        for words, bg, fg in ((red, "F4B6AD", "8A1C0E"), (amber, "FFE08A", "5A4300"), (green, "B7E1C4", "134D2B")):
            for w in words:
                ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first_cell}="{w}"'], fill=PatternFill("solid", bgColor=bg, fgColor=bg), font=Font(color=fg, bold=(bg == "F4B6AD"))))

    def status_workbook(fn, name, blurb, basis, areas, lines, group="Checklists"):
        GRP[0] = group
        wb = Workbook()
        sm = wb.active
        sm.title = "Summary"
        title(sm, name, "Work through each sheet. This page counts the results for you.")
        sm["A3"] = "Company name"
        sm["A3"].font = Font(bold=True)
        sm["B3"].fill = PatternFill("solid", fgColor=INPUT)
        sm["B3"].border = BOX
        sm["A4"] = "Date of the review"
        sm["A4"].font = Font(bold=True)
        sm["B4"].fill = PatternFill("solid", fgColor=INPUT)
        sm["B4"].border = BOX
        sm["B4"].number_format = DATEF
        head(sm, 6, ["Area", "Items", "Verified", "Gap found", "Not provided", "Not applicable", "Still to check"], [46, 10, 12, 12, 14, 16, 16])
        names = []
        for k, (sname, rows) in enumerate(areas):
            ws = wb.create_sheet(sname)
            names.append((sname, len(rows)))
            title(ws, f"{name}: {sname}", "Pick a status for each item. Record what you found in the last column.")
            head(ws, 4, ["Sr", "What to check", "Law or provision", "Documents to ask for", "Status", "Finding or remarks"], [5, 62, 30, 44, 16, 44])
            for i, (item, law, docs) in enumerate(rows):
                r = 5 + i
                ws.cell(row=r, column=1, value=i + 1)
                ws.cell(row=r, column=2, value=item)
                ws.cell(row=r, column=3, value=law)
                ws.cell(row=r, column=4, value=docs)
            last = 4 + len(rows) + 5
            style_rows(ws, 5, last, 6, formula_cols=(1,), wrap_cols=(2, 3, 4, 6))
            for r in range(5, 5 + len(rows)):
                for c in (2, 3, 4):
                    ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFFFFF")
            dv = DataValidation(type="list", formula1='"Verified,Gap found,Not provided,Not applicable"', allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"E5:E{last}")
            red_amber_green(ws, f"E5:E{last}", "E5")
            ws.freeze_panes = "C5"
            q = "'" + sname + "'!"
            r = 7 + k
            sm.cell(row=r, column=1, value=sname)
            sm.cell(row=r, column=2, value=f"=COUNTA({q}B5:B{4 + len(rows)})")
            for j, w in enumerate(("Verified", "Gap found", "Not provided", "Not applicable")):
                sm.cell(row=r, column=3 + j, value=f'=COUNTIF({q}E5:E{last},"{w}")')
            sm.cell(row=r, column=7, value=f"=B{r}-SUM(C{r}:F{r})")
        tr = 7 + len(areas)
        sm.cell(row=tr, column=1, value="Total").font = Font(bold=True)
        for c in range(2, 8):
            L_ = get_column_letter(c)
            sm.cell(row=tr, column=c, value=f"=SUM({L_}7:{L_}{tr - 1})").font = Font(bold=True)
        style_rows(sm, 7, tr, 7, formula_cols=tuple(range(1, 8)))
        red_amber_green(sm, f"D7:D{tr}", "D7", red=(), amber=(), green=())
        sm.conditional_formatting.add(f"D7:D{tr}", FormulaRule(formula=["D7>0"], fill=PatternFill("solid", bgColor="F4B6AD", fgColor="F4B6AD"), font=Font(bold=True, color="8A1C0E")))
        readme(wb, name, basis, lines)
        save(wb, fn, name, blurb, basis, ["Summary"] + [n for n, _ in names])
        return wb

    def attach_workbook(fn, name, blurb, basis, forms, lines):
        GRP[0] = "Checklists"
        wb = Workbook()
        ws = wb.active
        ws.title = "Inputs"
        title(ws, name, "Fill in the dates you know. Each form's sheet then shows its due date.")
        labels = [("Company name", None), ("Financial year ends on", DATEF), ("AGM date (planned or held)", DATEF), ("Date of the resolution or Board meeting", DATEF),
                  ("Date of appointment of a director, KMP or auditor", DATEF), ("Date of allotment of shares", DATEF), ("Date of incorporation", DATEF)]
        inputs(ws, labels, 1, 2, 3)
        ws.column_dimensions["A"].width = 50
        ws.column_dimensions["B"].width = 24
        ws["A11"] = "How it works"
        ws["A11"].font = Font(bold=True)
        ws["A12"] = "Each form has its own sheet. Pick Yes when the document is ready. The sheet counts what is still pending."
        for sname, due_formula, due_text, form_basis, rows in forms:
            s = wb.create_sheet(sname)
            title(s, f"{sname}: documents to attach and keep", form_basis)
            s["A3"] = "Due date"
            s["A3"].font = Font(bold=True)
            s["B3"] = due_formula if due_formula else "See the rule"
            s["B3"].number_format = DATEF
            s["B3"].fill = PatternFill("solid", fgColor=LIGHT)
            s["B3"].border = BOX
            s["C3"] = due_text
            s["C3"].font = Font(italic=True, color="4B5A70")
            s["A4"] = "Items still pending"
            s["A4"].font = Font(bold=True)
            n = len(rows)
            s["B4"] = f'=COUNTA(B7:B{6 + n})-COUNTIF(E7:E{6 + n + 5},"Yes")-COUNTIF(E7:E{6 + n + 5},"Not applicable")'
            s["B4"].fill = PatternFill("solid", fgColor=LIGHT)
            s["B4"].border = BOX
            head(s, 6, ["Sr", "Document or attachment", "When it is needed", "Prepared by or taken from", "Done?", "Remarks"], [5, 60, 36, 32, 14, 32])
            for i, (doc, when, who) in enumerate(rows):
                r = 7 + i
                s.cell(row=r, column=1, value=i + 1)
                s.cell(row=r, column=2, value=doc)
                s.cell(row=r, column=3, value=when)
                s.cell(row=r, column=4, value=who)
            last = 6 + n + 5
            style_rows(s, 7, last, 6, formula_cols=(1,), wrap_cols=(2, 3, 4, 6))
            for r in range(7, 7 + n):
                for c in (2, 3, 4):
                    s.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFFFFF")
            dv = DataValidation(type="list", formula1='"Yes,No,Not applicable"', allow_blank=True)
            s.add_data_validation(dv)
            dv.add(f"E7:E{last}")
            red_amber_green(s, f"E7:E{last}", "E7", red=("No",), amber=(), green=("Yes",))
            s.freeze_panes = "C7"
        readme(wb, name, basis, lines)
        save(wb, fn, name, blurb, basis, ["Inputs"] + [f[0] for f in forms])

    IN = "Inputs!"
    attach_workbook(
        "roc-filing-attachments-checklist.xlsx", "ROC filings: documents to attach", "For each annual and event-based ROC form: the documents to prepare, the due date, and a count of what is still pending.",
        "Companies Act, 2013 (sections 92, 117, 137, 139, 42, 12, 10A and others) and the Rules for each form",
        [
            ("AOC-4", f'=IF({IN}B5="","",{IN}B5+30)', "Within 30 days of the AGM (Section 137)", "Section 137; Rule 12, Companies (Accounts) Rules, 2014", [
                ("Financial statements: balance sheet, statement of profit and loss, cash flow statement (if required), statement of changes in equity and notes, signed", "Always", "Finance team; signed by directors and CFO/CS"),
                ("Auditor's report on the financial statements, with Annexures (CARO if it applies)", "Always", "Statutory auditor"),
                ("Board's report with its annexures (including the AOC-2, CSR report and the MGT-9 extract if the company has no website)", "Always", "Company Secretary; Board"),
                ("Consolidated financial statements and the auditor's report on them", "If the company has subsidiaries, associates or joint ventures (Section 129(3))", "Finance team; auditor"),
                ("Statement of subsidiaries, associates and joint ventures in Form AOC-1", "If it has them", "Finance team"),
                ("Secretarial audit report in Form MR-3", "Listed companies and the other classes in Rule 9", "Company Secretary in practice"),
                ("Comments of the Comptroller and Auditor-General", "Government companies", "CAG"),
                ("Financial statements in XBRL format", "Companies in the classes in Rule 3 of the XBRL Rules, 2015", "Finance team or consultant"),
                ("Digital signatures and certification by a professional in practice, as the e-form requires", "As the current e-form says", "Directors; practising professional"),
                ("Proof of fees and additional fees, if filed late", "If late", "Company Secretary"),
            ]),
            ("MGT-7 and MGT-7A", f'=IF({IN}B5="","",{IN}B5+60)', "Within 60 days of the AGM (Section 92)", "Section 92; Rule 11, Companies (Management and Administration) Rules, 2014", [
                ("List of members and debenture holders as on the date of the AGM or the last date", "As the form requires", "Company Secretary"),
                ("Details of the shares and debentures transferred since the last return", "Always", "Register of transfers"),
                ("Details of the directors and key managerial personnel, and the changes in them in the year", "Always", "Register of directors"),
                ("Remuneration of directors and key managerial personnel", "Always", "Finance team"),
                ("Penalties or punishment imposed on the company, its directors or officers, and the compounding", "If any", "Company Secretary"),
                ("Certificate of a Company Secretary in practice in Form MGT-8", "Listed companies and companies with paid-up capital of Rs. 10 crore or more or turnover of Rs. 50 crore or more (Section 92(2) and Rule 11)", "Company Secretary in practice"),
                ("Approval for the extension of the AGM, if given", "If the AGM was held late with the Registrar's extension", "Registrar's order"),
                ("Use Form MGT-7A, with no MGT-8, for One Person Companies and small companies", "If the company is one", "Company Secretary"),
            ]),
            ("ADT-1", f'=IF({IN}B7="","",{IN}B7+15)', "Within 15 days of the appointment (Section 139; Rule 4)", "Section 139; Rule 4, Companies (Audit and Auditors) Rules, 2014", [
                ("Certified copy of the Board or members' resolution appointing the auditor", "Always", "Company Secretary"),
                ("Written consent of the auditor", "Always", "Auditor"),
                ("Certificate from the auditor that the appointment satisfies Section 141", "Always", "Auditor"),
                ("Copy of the intimation sent by the company to the auditor", "Always", "Company Secretary"),
                ("Peer review certificate", "Auditors of listed companies, as SEBI requires", "Auditor"),
                ("Notice of the AGM and the explanatory statement", "If the appointment was made at the AGM", "Company Secretary"),
            ]),
            ("DPT-3", f'=IF({IN}B4="","",DATE(YEAR({IN}B4),6,30))', "On or before 30 June each year, with data as on 31 March (Rule 16 and 16A)", "Sections 73 to 76 and 179; Rule 16 and 16A, Companies (Acceptance of Deposits) Rules, 2014", [
                ("Details of deposits outstanding on 31 March", "If the company has any", "Finance team"),
                ("Details of money received that is not a deposit (Rule 2(1)(c)), such as loans from directors with the declaration of the source of funds", "Always", "Finance team"),
                ("Audited balance sheet for the year", "Always", "Finance team"),
                ("Auditor's certificate or report on the particulars", "As the form requires", "Statutory auditor"),
                ("Declaration of the director from whom money is taken, that it is not out of borrowed funds", "If money was received from directors", "Directors"),
            ]),
            ("MGT-14", f'=IF({IN}B6="","",{IN}B6+30)', "Within 30 days of passing the resolution (Section 117)", "Section 117; Rule 24, Companies (Management and Administration) Rules, 2014", [
                ("Certified true copy of the resolution", "Always", "Company Secretary"),
                ("Notice of the meeting with the explanatory statement", "If passed at a meeting", "Company Secretary"),
                ("Altered memorandum or articles of association", "If the resolution alters them", "Company Secretary"),
                ("Copy of the Board resolution that approved the proposal", "Where the Act needs the Board's approval first", "Company Secretary"),
                ("Draft agreement or papers the resolution approves", "If any", "Company"),
                ("Valuation report", "Where the Act needs it, such as a preferential allotment", "Registered valuer"),
            ]),
            ("DIR-12", f'=IF({IN}B7="","",{IN}B7+30)', "Within 30 days of the appointment, change or cessation (Section 170(2))", "Section 170(2); Rule 18, Companies (Appointment and Qualification of Directors) Rules, 2014", [
                ("Letter of appointment, and proof of the Board or members' resolution", "For an appointment", "Company Secretary"),
                ("Consent of the director in Form DIR-2", "For a director's appointment", "Director"),
                ("Notice of interest in Form MBP-1", "For a director's appointment", "Director"),
                ("Declaration in Form DIR-8 that the director is not disqualified", "For a director's appointment", "Director"),
                ("Declaration of independence under Section 149(7)", "For an independent director", "Director"),
                ("Resignation letter, with proof that it was received", "For a resignation", "Director"),
                ("Form DIR-11 filed by the director, if the director has filed it", "For a resignation", "Director"),
                ("Letter of appointment or resignation of a key managerial personnel, with the Board resolution", "For key managerial personnel", "Company Secretary"),
            ]),
            ("PAS-3", f'=IF({IN}B8="","",{IN}B8+15)', "Within 15 days of a private placement or preferential allotment. For other allotments, within 30 days (Section 39(4))", "Sections 39 and 42; Rule 12 and 14, Companies (Prospectus and Allotment of Securities) Rules, 2014", [
                ("List of allottees with name, address, number of shares, amount and the date of allotment", "Always", "Company Secretary"),
                ("Board resolution allotting the shares", "Always", "Company Secretary"),
                ("Special resolution approving the offer, and the MGT-14 receipt", "For a private placement or preferential allotment", "Company Secretary"),
                ("Valuation report of a registered valuer", "For a preferential allotment, and a private placement where the Rules need it", "Registered valuer"),
                ("Copies of the offer letters in Form PAS-4 and the record in Form PAS-5", "For a private placement", "Company Secretary"),
                ("Copy of the contract and valuation, where shares are issued for consideration other than cash", "If any", "Company"),
                ("Proof of receipt of the money from the investor's own bank account", "Keep with the file", "Bank statement"),
            ]),
            ("INC-20A", f'=IF({IN}B9="","",{IN}B9+180)', "Within 180 days of incorporation (Section 10A)", "Section 10A, Companies Act, 2013; Rule 23A, Companies (Incorporation) Rules, 2014", [
                ("Declaration by a director that every subscriber has paid the value of the shares agreed to be taken by him", "Always", "Director"),
                ("Bank statement or bank certificate showing the receipt of the subscription money from every subscriber", "Always", "Bank"),
                ("Proof that the registered office is verified, with the INC-22 receipt", "Always", "Company Secretary"),
                ("Certificate of a practising professional, if the e-form needs it", "As the current e-form says", "Practising professional"),
            ]),
        ],
        ["1. Fill in the dates on the Inputs sheet that apply to the filing you are working on.", "2. Open the sheet for the form. The due date is shown at the top. Pick Yes against each document when it is ready.", "3. The lists come from the form and the Rules. MCA changes its e-forms often, so check the attachment tab of the current e-form before you file.", "4. The due dates are the dates the Act gives. Check whether any extension or exemption applies."])

    DD = [
        ("Incorporation and constitution", [
            ("Certificate of incorporation, and any fresh certificates after a change of name or conversion", "Sections 7 and 13", "Certificates, MCA master data"),
            ("Memorandum and articles of association, with every alteration, and the filing of each alteration", "Sections 4, 5, 13, 14", "MOA and AOA, MGT-14 receipts, INC-24 where the name changed"),
            ("Objects clause covers the business actually carried on", "Section 4(1)(c)", "MOA, business description"),
            ("Registered office, its proof, and the notice of any change (INC-22)", "Section 12", "INC-22, utility bill, lease deed"),
            ("Commencement of business declaration (INC-20A) for a company incorporated after 2019", "Section 10A", "INC-20A receipt"),
            ("Company's name, display of name and registered office on the premises and on its documents", "Section 12(3)", "Letterheads, board outside the office"),
            ("Status as a private, public or small company, and the conditions for that status", "Sections 2(68), 2(71), 2(85)", "Articles; financial statements"),
        ]),
        ("Share capital and members", [
            ("Authorised, issued, subscribed and paid-up capital agree to the register of members and to the financial statements", "Sections 2(8), 61", "Register of members, balance sheet"),
            ("Each allotment: Board resolution, special resolution if needed, PAS-3, valuation, and money received through a bank", "Sections 39, 42, 62", "Allotment papers"),
            ("Share certificates issued within two months, and the register of renewed and duplicate certificates", "Sections 46, 56", "Certificates, Form SH-2 register"),
            ("Each transfer and transmission: signed instrument SH-4, stamp duty, Board approval, and entry in the register", "Section 56", "SH-4, Board minutes"),
            ("Register of members and index of members, with the entries up to date", "Section 88", "Register in Form MGT-1"),
            ("Pledges, lien, or encumbrances on shares, and their records", "Section 124 and the articles", "Pledge deeds, depository records"),
            ("Dematerialisation of securities (unlisted public companies, and where the Rules require)", "Sections 29, 2(68); Rule 9B", "Depository statement"),
            ("Significant beneficial owners: declarations and the register", "Section 90", "BEN-1, BEN-2, BEN-3 register"),
            ("Shareholders' agreement and its consistency with the articles", "Section 5", "SHA, articles"),
            ("Sweat equity, ESOPs and buy-back: approvals, registers and filings", "Sections 54, 62(1)(b), 68", "SH-3, SH-6, SH-10 registers"),
        ]),
        ("Directors and KMP", [
            ("Each director's DIN, KYC, consent in DIR-2 and declaration in DIR-8", "Sections 152, 164; Rules 12 and 14", "DIN allotment, DIR-2, DIR-8"),
            ("Number of directors, resident director, woman director and independent directors, where required", "Sections 149, 3(1), 152", "Board list, articles"),
            ("Appointment and reappointment of each director, and their approvals; retirement by rotation", "Sections 152, 161, 196", "Board and members' resolutions"),
            ("Disqualifications and the checks made on the DIN holders", "Section 164", "DIR-8, MCA data"),
            ("Disclosures of interest in MBP-1, and the register", "Section 184", "MBP-1, register"),
            ("Key managerial personnel appointed as required, and DIR-12 filed", "Section 203", "Appointment letters, DIR-12"),
            ("Remuneration paid to directors and KMP within the limits, with the approvals", "Sections 196, 197, Schedule V", "Resolutions, payroll"),
            ("Directors' and officers' insurance, and loans to directors", "Sections 185 and 197(13)", "Policies, ledgers"),
            ("Number of directorships held by each director", "Section 165", "Director list"),
        ]),
        ("Meetings and minutes", [
            ("At least four Board meetings a year, with not more than 120 days between two meetings", "Section 173", "Notices, minutes"),
            ("Notice, agenda and notes sent to all directors on time, and quorum present", "Sections 173, 174; SS-1", "Notices, attendance"),
            ("Minutes signed by the chairman within thirty days, and pages numbered and kept in the minute book", "Section 118; SS-1", "Minute books"),
            ("AGM held within six months of the year end, with not more than fifteen months between two AGMs", "Section 96", "AGM notices, minutes"),
            ("Notice of general meetings, with the explanatory statements and proxies", "Sections 101, 102, 105", "Notices, proxies, dispatch proofs"),
            ("Resolutions that need a special resolution were passed as such, and MGT-14 filed", "Sections 114, 117", "Minutes, MGT-14"),
            ("Committees constituted as the Act requires, and their meetings", "Sections 177, 178, 135", "Committee minutes"),
            ("Resolutions by circulation and the matters that may not be passed that way", "Section 175; Rule 8", "Circular resolutions"),
            ("Postal ballot and e-voting, where required", "Sections 108, 110", "Scrutinizer's reports"),
        ]),
        ("Filings with the ROC", [
            ("Annual financial statements: AOC-4 filed on time", "Section 137", "AOC-4 receipt"),
            ("Annual return: MGT-7 or MGT-7A filed on time", "Section 92", "MGT-7 receipt"),
            ("Auditor's appointment: ADT-1", "Section 139", "ADT-1 receipt"),
            ("Return of deposits, DPT-3", "Rule 16 and 16A", "DPT-3 receipts"),
            ("MSME-1 returns for payments to micro and small enterprises delayed over 45 days", "Section 405", "MSME-1 receipts"),
            ("Return of allotment PAS-3, and other event forms (MGT-14, DIR-12, INC-22, CHG-1)", "Various", "SRNs and receipts"),
            ("Additional fees paid for any late filing, and notices received from the Registrar", "Section 403", "Challans, notices"),
            ("Status of the company on the MCA portal (active compliant), and any defaults or strike-off risk", "Section 248", "MCA master data"),
        ]),
        ("Charges and borrowings", [
            ("Each charge is registered with the ROC (CHG-1) within the time allowed, and satisfaction filed (CHG-4)", "Sections 77 to 87", "CHG-1, CHG-4"),
            ("Register of charges matches the charges on the MCA records and the lenders' records", "Section 85", "Register of charges"),
            ("Borrowing within the limit approved by the members", "Section 180(1)(c)", "Special resolution, financials"),
            ("Mortgage or sale of undertaking approved by special resolution where it applies", "Section 180(1)(a)", "Resolution"),
            ("Loans, guarantees and investments within the limits of Section 186, and the register", "Sections 185 and 186", "Resolutions, MBP-2 register"),
            ("Loan agreements and the compliance with their covenants", "Contract", "Loan documents"),
        ]),
        ("Related parties and others", [
            ("Related party transactions: Board approval, members' approval, AOC-2, and arm's length proof", "Section 188", "RPT policy, resolutions, AOC-2"),
            ("Register of contracts with related parties", "Section 189", "Register in MBP-4"),
            ("CSR: committee, policy, spend and the report", "Section 135", "CSR policy, reports"),
            ("Whistle-blower policy and the audit committee, where required", "Section 177", "Policy, minutes"),
            ("Pending litigation, notices from regulators, and their disclosure", "Various", "Legal register"),
            ("Insurance, licences, registrations under other laws and their renewals", "Various", "Licences, certificates"),
            ("Statutory registers kept at the registered office and open for inspection", "Section 94", "Registers"),
        ]),
        ("Listed company items", [
            ("Shareholding pattern, corporate governance report and investor complaints filed each quarter", "Regulations 31, 27, 13", "Exchange filings"),
            ("Financial results filed within 45 days of each quarter and 60 days of the year end", "Regulation 33", "Exchange filings"),
            ("Disclosure of events under Regulation 30 within the time allowed", "Regulation 30", "Exchange filings, policy"),
            ("Board and committee composition: independent directors, woman director, chairman", "Regulations 17, 18, 19, 20", "Board list"),
            ("Annual secretarial compliance report and secretarial audit report", "Regulation 24A", "Reports"),
            ("Insider trading code, structured digital database and trading window", "PIT Regulations, 2015", "Code, SDD"),
            ("Promoter and public shareholding, and promoters' pledges", "SAST and LODR", "Disclosures"),
            ("Website disclosures required by Regulation 46", "Regulation 46", "Website"),
        ]),
    ]
    status_workbook("secretarial-due-diligence-checklist.xlsx", "Secretarial due diligence checklist", "A checklist for reviewing a company's secretarial position before an investment, a deal or a lender's review.",
                    "Companies Act, 2013 and the Rules; SEBI (LODR) Regulations, 2015 for listed companies",
                    DD, ["1. Fill in the company name and the date on the Summary sheet.", "2. Open each area's sheet. Pick Verified, Gap found, Not provided or Not applicable against each item, and write what you found.", "3. The Summary sheet counts the results and shows the gaps in red.", "4. This is a working list, not a complete one. Add the items that your deal or your client needs."])

    MR = [
        ("Companies Act, 2013", [
            ("Constitution of the Board, including the independent directors, woman director and resident director", "Sections 149, 152; Rule 3 and 4", "Board list, DIN, declarations"),
            ("Appointment, reappointment and cessation of directors, and key managerial personnel, with DIR-12", "Sections 152, 161, 170, 203", "Resolutions, DIR-12"),
            ("Board and committee meetings: notice, agenda, quorum, gap between meetings and minutes", "Sections 173 to 175, 118", "Minutes, notices, attendance"),
            ("General meetings, e-voting and postal ballot, and filing of the resolutions", "Sections 96, 101, 108, 110, 117", "Notices, minutes, scrutinizer's reports"),
            ("Statutory registers kept and updated", "Sections 85, 88, 170, 186, 187, 189", "Registers"),
            ("Disclosures of interest and declarations by directors", "Sections 164, 184", "MBP-1, DIR-8"),
            ("Related party transactions, loans, guarantees and investments", "Sections 185, 186, 188", "Resolutions, AOC-2"),
            ("Share capital: allotments, transfers, certificates, buy-back, sweat equity and ESOPs", "Sections 39, 42, 46, 54, 56, 62, 68", "PAS-3, SH forms"),
            ("Annual filings: AOC-4, MGT-7, ADT-1, DPT-3, MSME-1 and their timeliness", "Sections 92, 137, 139; Rule 16", "Receipts"),
            ("Dividend, unclaimed dividend, and transfers to the IEPF", "Sections 123 to 125", "Dividend records, IEPF-1"),
            ("Board's report content and its disclosures", "Section 134", "Board's report"),
            ("CSR, nomination and remuneration policy, vigil mechanism and audit committee, where they apply", "Sections 135, 177, 178", "Policies, minutes"),
            ("Charges registered and satisfied", "Sections 77 to 87", "CHG forms"),
            ("Deposits and money received that are not deposits", "Sections 73 to 76", "DPT-3, ledgers"),
        ]),
        ("SCRA and Depositories", [
            ("Compliance with the Securities Contracts (Regulation) Act, 1956 and the Rules", "SCRA, 1956", "Listing records"),
            ("Dematerialisation, and compliance with the Depositories Act and the bye-laws", "Depositories Act, 1996", "Depository statements"),
            ("Reconciliation of share capital audit report, where it applies", "SEBI regulations", "Reports"),
        ]),
        ("FEMA", [
            ("Foreign direct investment: reporting of inflows (FC-GPR and FC-TRS) and the pricing", "FEMA, 1999; NDI Rules, 2019", "FC-GPR, FC-TRS"),
            ("Overseas direct investment and the reporting", "FEMA; ODI rules", "Forms"),
            ("External commercial borrowings and their reporting", "FEMA; ECB framework", "Forms"),
            ("Annual return on foreign liabilities and assets (FLA)", "RBI", "FLA return"),
        ]),
        ("SEBI laws (listed company)", [
            ("SAST: disclosures by acquirers and promoters", "Regulations 29 to 31", "Disclosures"),
            ("Insider trading: code, SDD and trading window closure", "PIT Regulations, 2015", "Code, SDD"),
            ("ICDR: issue of capital, if any, and the disclosures", "ICDR Regulations, 2018", "Offer documents"),
            ("Share-based employee benefits and sweat equity", "SBEB and SE Regulations, 2021", "Resolutions, registers"),
            ("Issue and listing of non-convertible securities", "NCS Regulations, 2021", "Records"),
            ("Registrar to an issue and share transfer agent", "RTA Regulations, 1993", "RTA agreement"),
            ("Delisting and buy-back, if any", "Delisting Regulations, 2021; Buy-back Regulations, 2018", "Records"),
            ("LODR: Board composition, committees, quarterly filings, disclosures and website", "LODR Regulations, 2015", "Filings"),
        ]),
        ("Secretarial Standards", [
            ("SS-1: Board and committee meetings, notices, minutes and records", "Secretarial Standard 1", "Minute books"),
            ("SS-2: General meetings, notices, voting, minutes and records", "Secretarial Standard 2", "Minute books"),
            ("Dividend (SS-3) and Board's report (SS-4), where they apply", "Secretarial Standards 3 and 4", "Records"),
        ]),
        ("Other applicable laws", [
            ("Laws specifically applicable to the business: list them with the management", "Industry laws", "Management representation"),
            ("Labour laws: registrations, returns and payments (provident fund, ESI, minimum wages, bonus, gratuity)", "Labour laws", "Challans, returns"),
            ("Environment and pollution control, and factory laws, if they apply", "Environment laws", "Licences, consents"),
            ("Tax filings status (income tax, GST, TDS), as the management reports it", "Tax laws", "Management report"),
            ("Other sector rules (for example RBI, IRDAI, TRAI, FSSAI)", "Sector laws", "Licences"),
        ]),
    ]
    status_workbook("secretarial-audit-checklist-mr3.xlsx", "Secretarial audit checklist (Form MR-3)", "A working checklist for a secretarial audit, arranged in the order of the laws that Form MR-3 lists.",
                    "Section 204, Companies Act, 2013; Form MR-3",
                    MR, ["1. Fill in the company and the audit date on the Summary sheet.", "2. Work through each law's sheet. Record each item as Verified, Gap found, Not provided or Not applicable.", "3. Each Gap found goes into the observations paragraph of the MR-3 report. The Summary sheet counts them.", "4. Remove the laws that do not apply to the company, and add the ones that do."])

    # ---- accounts and finance ---------------------------------------------------------
    GRP[0] = "Accounts and finance"
    LIVES = [
        ("Buildings (other than factory buildings): RCC frame structure", 60, "Schedule II, Part C"),
        ("Buildings (other than factory buildings): other than RCC frame structure", 30, "Schedule II, Part C"),
        ("Factory buildings", 30, "Schedule II, Part C"),
        ("Fences, wells and tube wells", 5, "Schedule II, Part C"),
        ("Other buildings, including temporary structures", 3, "Schedule II, Part C"),
        ("Roads: carpet roads (RCC)", 10, "Schedule II, Part C"),
        ("Roads: carpet roads (other than RCC)", 5, "Schedule II, Part C"),
        ("Roads: non-carpet roads", 3, "Schedule II, Part C"),
        ("Bridges, culverts and bunders", 30, "Schedule II, Part C"),
        ("Plant and machinery: general (single shift)", 15, "Schedule II, Part C"),
        ("Furniture and fittings: general", 10, "Schedule II, Part C"),
        ("Furniture and fittings: hotels, restaurants, boarding houses, schools, hospitals, hostels", 8, "Schedule II, Part C"),
        ("Motor cycles, scooters and other mopeds", 10, "Schedule II, Part C"),
        ("Motor buses, lorries, cars and taxis used in the business of running them on hire", 6, "Schedule II, Part C"),
        ("Motor buses, lorries and cars other than those used on hire", 8, "Schedule II, Part C"),
        ("Electrical installations and equipment", 10, "Schedule II, Part C"),
        ("Office equipment: general", 5, "Schedule II, Part C"),
        ("Computers: servers and networks", 6, "Schedule II, Part C"),
        ("Computers: end-user devices such as desktops and laptops", 3, "Schedule II, Part C"),
        ("Aircraft: aeroplanes and helicopters", 20, "Schedule II, Part C"),
    ]

    def lives_sheet(wb):
        ws = wb.create_sheet("Useful lives")
        title(ws, "Useful lives under Schedule II", "Add or change a row if your class is not here. The calculator picks the life from this list.")
        head(ws, 4, ["Asset class", "Useful life (years)", "Reference"], [78, 20, 28])
        for i, (c, y, ref) in enumerate(LIVES):
            ws.cell(row=5 + i, column=1, value=c)
            ws.cell(row=5 + i, column=2, value=y)
            ws.cell(row=5 + i, column=3, value=ref)
        style_rows(ws, 5, 40, 3)
        ws.cell(row=42, column=1, value="Schedule II also gives lives for ships, plant and machinery in special industries (such as telecom, power, sugar), medical equipment and others. Add them here. Check every figure against the current Schedule II. A company may use a different life if it gives reasons.").alignment = Alignment(wrap_text=True)
        ws.row_dimensions[42].height = 48
        return ws

    # depreciation calculator
    wb = Workbook()
    ws = wb.active
    ws.title = "Calculator"
    title(ws, "Depreciation calculator (Schedule II, Companies Act, 2013)", "One asset, one year. Pick the class, fill the yellow cells, and read the depreciation for the year.")
    ws.column_dimensions["A"].width = 56
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 60
    inputs(ws, [("Company name", None), ("Financial year starts on", DATEF)], 1, 2, 3)
    ws["A5"] = "Financial year ends on"
    ws["B5"] = '=IF(B4="","",EDATE(B4,12)-1)'
    ws["A6"] = "Days in the year"
    ws["B6"] = '=IF(B4="","",B5-B4+1)'
    for r in (5, 6):
        ws.cell(row=r, column=1).font = Font(bold=True)
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row=r, column=2).border = BOX
    ws["B5"].number_format = DATEF
    inputs(ws, [("Asset description", None), ("Asset class (pick from the list)", None), ("Date put to use", DATEF), ("Cost (Rs.)", "#,##0.00"), ("Residual value as a share of cost", "0.00%"),
                ("Useful life in years (leave blank to use the list)", "0"), ("Method", None), ("Depreciation already charged before this year (Rs.)", "#,##0.00")], 1, 2, 8)
    ws["B12"] = 0.05
    ws["B14"] = "SLM"
    dv1 = DataValidation(type="list", formula1="='Useful lives'!$A$5:$A$40", allow_blank=True)
    dv2 = DataValidation(type="list", formula1='"SLM,WDV"', allow_blank=False)
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    dv1.add("B9")
    dv2.add("B14")
    calc = [("Useful life used (years)", '=IF(B13<>"",B13,IFERROR(VLOOKUP(B9,\'Useful lives\'!$A$5:$B$40,2,FALSE),""))', "0.0"),
            ("Residual value (Rs.)", '=IF(B11="","",B11*B12)', "#,##0.00"),
            ("Depreciable amount (Rs.)", '=IF(B11="","",B11-B18)', "#,##0.00"),
            ("Days the asset was in use this year", '=IF(OR(B4="",B10=""),"",MAX(0,B5-MAX(B4,B10)+1))', "0"),
            ("Written-down value at the start of the year (Rs.)", '=IF(B11="","",B11-N(B15))', "#,##0.00"),
            ("Depreciation for the year (Rs.)", '=IF(OR(B11="",B17="",B20="",B6=""),"",IF(B14="WDV",MAX(0,MIN(B21-B18,B21*(1-(B18/B11)^(1/B17))*B20/B6)),MAX(0,MIN(B19-N(B15),B19/B17*B20/B6))))', "#,##0.00"),
            ("Written-down value at the end of the year (Rs.)", '=IF(B22="","",B21-B22)', "#,##0.00"),
            ("Total depreciation charged to date (Rs.)", '=IF(B22="","",N(B15)+B22)', "#,##0.00")]
    ws["A16"] = "Worked out for you"
    ws["A16"].font = Font(bold=True, color=NAVY, size=12)
    for i, (lab, f, fmt) in enumerate(calc):
        r = 17 + i
        ws.cell(row=r, column=1, value=lab).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=f)
        c.number_format = fmt
        c.fill = PatternFill("solid", fgColor=LIGHT)
        c.border = BOX
    ws["C22"] = "Straight line: (cost less residual value) / life, in proportion to the days used. Written-down value: the opening value times a rate of 1 less (residual / cost) to the power of 1 / life; this needs a residual value above zero."
    ws["C22"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[22].height = 62
    lives_sheet(wb)
    readme(wb, "Depreciation calculator (Schedule II)", "Schedule II, Companies Act, 2013 (useful lives and residual value of up to five per cent of the cost)",
           ["1. Type the start of the financial year (for example 01-Apr-2026). The end of the year is worked out.", "2. Pick the asset class, type the date it was put to use and the cost. Leave the residual value at 5% unless the company has a reason to change it.", "3. Read the depreciation for the year and the closing value. For an asset bought during the year, the depreciation is in proportion to the days.", "4. A company may use a different life or method if it can justify it and discloses it. Overwrite the life in the cell if you do."])
    save(wb, "depreciation-calculator-schedule-ii.xlsx", "Depreciation calculator (Schedule II)", "Depreciation of one asset for a year, using the useful lives in Schedule II, by the straight-line or the written-down-value method.", "Schedule II, Companies Act, 2013", ["Calculator", "Useful lives"])

    # fixed asset register
    wb = Workbook()
    ws = wb.active
    ws.title = "Register"
    title(ws, "Fixed asset register with Schedule II depreciation", "Enter each asset once. The depreciation for the year is worked out from the lives on the 'Useful lives' sheet.")
    ws["A3"] = "Financial year starts on"
    ws["A3"].font = Font(bold=True)
    ws["B3"].fill = PatternFill("solid", fgColor=INPUT)
    ws["B3"].border = BOX
    ws["B3"].number_format = DATEF
    ws["A4"] = "Financial year ends on"
    ws["A4"].font = Font(bold=True)
    ws["B4"] = '=IF(B3="","",EDATE(B3,12)-1)'
    ws["A5"] = "Days in the year"
    ws["A5"].font = Font(bold=True)
    ws["B5"] = '=IF(B3="","",B4-B3+1)'
    for r in (4, 5):
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row=r, column=2).border = BOX
    ws["B4"].number_format = DATEF
    FR_H = ["Asset ID", "Description", "Asset class", "Date put to use", "Cost (Rs.)", "Residual value (share of cost)", "Life override (years)", "Life used (years)", "Residual value (Rs.)",
            "Depreciation charged before this year (Rs.)", "Date of disposal", "Days used this year", "Depreciation for the year (Rs.)", "Total depreciation to date (Rs.)", "Written-down value (Rs.)", "Remarks"]
    head(ws, 8, FR_H, [11, 34, 40, 14, 16, 14, 12, 11, 16, 20, 14, 11, 18, 18, 18, 28])
    first, lastr = 9, 308
    for r in range(first, lastr + 1):
        ws.cell(row=r, column=6, value=0.05)
        ws.cell(row=r, column=8, value=f'=IF(C{r}="","",IF(G{r}<>"",G{r},IFERROR(VLOOKUP(C{r},\'Useful lives\'!$A$5:$B$40,2,FALSE),"")))')
        ws.cell(row=r, column=9, value=f'=IF(E{r}="","",E{r}*IF(F{r}="",0.05,F{r}))')
        ws.cell(row=r, column=12, value=f'=IF(OR(D{r}="",E{r}="",$B$3=""),"",MAX(0,MIN($B$4,IF(K{r}="",$B$4,K{r}))-MAX($B$3,D{r})+1))')
        ws.cell(row=r, column=13, value=f'=IF(OR(L{r}="",H{r}="",H{r}=0),"",MAX(0,MIN(E{r}-I{r}-N(J{r}),(E{r}-I{r})/H{r}*L{r}/$B$5)))')
        ws.cell(row=r, column=14, value=f'=IF(M{r}="","",N(J{r})+M{r})')
        ws.cell(row=r, column=15, value=f'=IF(N{r}="","",E{r}-N{r})')
    style_rows(ws, first, lastr, 16, formula_cols=(8, 9, 12, 13, 14, 15), date_cols=(4, 11), num_cols=(5, 9, 10, 13, 14, 15), wrap_cols=(2, 16))
    for r in range(first, lastr + 1):
        ws.cell(row=r, column=6).number_format = "0.00%"
    ws["A6"] = "Totals"
    ws["A6"].font = Font(bold=True)
    for col in ("E", "I", "J", "M", "N", "O"):
        c = ws[f"{col}6"]
        c.value = f"=SUM({col}{first}:{col}{lastr})"
        c.number_format = "#,##0.00"
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor=LIGHT)
        c.border = BOX
    dvc = DataValidation(type="list", formula1="='Useful lives'!$A$5:$A$40", allow_blank=True)
    ws.add_data_validation(dvc)
    dvc.add(f"C{first}:C{lastr}")
    ws.freeze_panes = "C9"
    lives_sheet(wb)
    readme(wb, "Fixed asset register (Schedule II)", "Schedule II, Companies Act, 2013; Section 123(2)",
           ["1. Type the start of the financial year in cell B3.", "2. List each asset: its ID, class, date put to use and cost. For assets bought in earlier years, type the depreciation charged before this year.", "3. The depreciation for the year, the total to date and the written-down value come out. Assets bought during the year get depreciation in proportion to the days.", "4. If you dispose of an asset, type the date. Depreciation stops on that date. Show the sale and the profit or loss in your accounts separately: this file does not."])
    save(wb, "fixed-asset-register-schedule-ii.xlsx", "Fixed asset register (Schedule II)", "A register of up to 300 assets with the depreciation worked out under Schedule II, and the written-down value.", "Schedule II, Companies Act, 2013", ["Register", "Useful lives"])

    # Schedule III ratios
    wb = Workbook()
    wi = wb.active
    wi.title = "Inputs"
    title(wi, "Schedule III ratios: your figures", "Type the amounts from the financial statements, in the same unit, for the current year (CY) and the previous year (PY).")
    wi.column_dimensions["A"].width = 66
    for c in "BCD":
        wi.column_dimensions[c].width = 22
    wi["A3"] = "Company name"
    wi["A3"].font = Font(bold=True)
    wi["B3"].fill = PatternFill("solid", fgColor=INPUT)
    wi["B3"].border = BOX
    head(wi, 4, ["Balance sheet items (closing)", "As at the end of CY", "As at the end of PY", "As at the end of the year before PY"], [66, 22, 22, 22])
    BS = ["Current assets", "Current liabilities", "Total debt (borrowings and lease liabilities)", "Shareholders' equity", "Inventories", "Trade receivables", "Trade payables",
          "Tangible net worth", "Deferred tax liability", "Investments (funds in treasury investments)"]
    for i, lab in enumerate(BS):
        wi.cell(row=5 + i, column=1, value=lab).font = Font(bold=True)
    style_rows(wi, 5, 14, 4, num_cols=(2, 3, 4))
    for r in range(5, 15):
        wi.cell(row=r, column=1).fill = PatternFill("solid", fgColor="FFFFFF")
    head(wi, 16, ["Profit and loss items (for the year)", "CY", "PY", ""], [66, 22, 22, 22])
    PL = ["Revenue from operations (net sales)", "Net credit sales", "Cost of goods sold", "Net credit purchases", "Net profit after tax", "Depreciation and other amortisation", "Finance cost (interest)",
          "Other non-cash adjustments", "Principal repayments of borrowings in the year", "Lease payments in the year", "Profit before tax", "Income from invested funds (treasury)"]
    for i, lab in enumerate(PL):
        wi.cell(row=17 + i, column=1, value=lab).font = Font(bold=True)
    style_rows(wi, 17, 28, 3, num_cols=(2, 3))
    for r in range(17, 29):
        wi.cell(row=r, column=1).fill = PatternFill("solid", fgColor="FFFFFF")
    wr = wb.create_sheet("Ratios")
    title(wr, "Schedule III ratios", "Worked out from the Inputs sheet. Schedule III needs a reason where a ratio changes by more than 25% from the previous year.")
    head(wr, 4, ["Ratio", "Numerator", "Denominator", "CY", "PY", "Change from PY", "Reason, if the change is more than 25%"], [34, 44, 44, 14, 14, 16, 60])
    I_ = "Inputs!"

    def sf(num, den):
        return f'=IFERROR(({num})/({den}),"")'
    ratios = [
        ("Current ratio", "Current assets", "Current liabilities", sf(f"{I_}B5", f"{I_}B6"), sf(f"{I_}C5", f"{I_}C6"), "0.00"),
        ("Debt-equity ratio", "Total debt", "Shareholders' equity", sf(f"{I_}B7", f"{I_}B8"), sf(f"{I_}C7", f"{I_}C8"), "0.00"),
        ("Debt service coverage ratio", "Net profit after tax + depreciation + finance cost + other non-cash adjustments", "Finance cost + lease payments + principal repayments",
         sf(f"{I_}B21+{I_}B22+{I_}B23+{I_}B24", f"{I_}B23+{I_}B26+{I_}B25"), sf(f"{I_}C21+{I_}C22+{I_}C23+{I_}C24", f"{I_}C23+{I_}C26+{I_}C25"), "0.00"),
        ("Return on equity ratio", "Net profit after tax", "Average shareholders' equity", sf(f"{I_}B21", f"({I_}B8+{I_}C8)/2"), sf(f"{I_}C21", f"({I_}C8+{I_}D8)/2"), "0.00%"),
        ("Inventory turnover ratio", "Cost of goods sold", "Average inventories", sf(f"{I_}B19", f"({I_}B9+{I_}C9)/2"), sf(f"{I_}C19", f"({I_}C9+{I_}D9)/2"), "0.00"),
        ("Trade receivables turnover ratio", "Net credit sales", "Average trade receivables", sf(f"{I_}B18", f"({I_}B10+{I_}C10)/2"), sf(f"{I_}C18", f"({I_}C10+{I_}D10)/2"), "0.00"),
        ("Trade payables turnover ratio", "Net credit purchases", "Average trade payables", sf(f"{I_}B20", f"({I_}B11+{I_}C11)/2"), sf(f"{I_}C20", f"({I_}C11+{I_}D11)/2"), "0.00"),
        ("Net capital turnover ratio", "Revenue from operations", "Working capital (current assets less current liabilities)", sf(f"{I_}B17", f"{I_}B5-{I_}B6"), sf(f"{I_}C17", f"{I_}C5-{I_}C6"), "0.00"),
        ("Net profit ratio", "Net profit after tax", "Revenue from operations", sf(f"{I_}B21", f"{I_}B17"), sf(f"{I_}C21", f"{I_}C17"), "0.00%"),
        ("Return on capital employed", "Profit before tax + finance cost", "Tangible net worth + total debt + deferred tax liability", sf(f"{I_}B27+{I_}B23", f"{I_}B12+{I_}B7+{I_}B13"), sf(f"{I_}C27+{I_}C23", f"{I_}C12+{I_}C7+{I_}C13"), "0.00%"),
        ("Return on investment", "Income from invested funds", "Average invested funds", sf(f"{I_}B28", f"({I_}B14+{I_}C14)/2"), sf(f"{I_}C28", f"({I_}C14+{I_}D14)/2"), "0.00%"),
    ]
    for i, (nm, nu, de, fcy, fpy, fmt) in enumerate(ratios):
        r = 5 + i
        wr.cell(row=r, column=1, value=nm)
        wr.cell(row=r, column=2, value=nu)
        wr.cell(row=r, column=3, value=de)
        wr.cell(row=r, column=4, value=fcy).number_format = fmt
        wr.cell(row=r, column=5, value=fpy).number_format = fmt
        wr.cell(row=r, column=6, value=f'=IF(OR(D{r}="",E{r}="",E{r}=0),"",(D{r}-E{r})/ABS(E{r}))').number_format = "0.0%"
    style_rows(wr, 5, 15, 7, formula_cols=(1, 2, 3, 4, 5, 6), wrap_cols=(2, 3, 7))
    for r in range(5, 16):
        wr.cell(row=r, column=7).fill = PatternFill("solid", fgColor=INPUT)
        wr.cell(row=r, column=4).number_format = ratios[r - 5][5]
        wr.cell(row=r, column=5).number_format = ratios[r - 5][5]
        wr.cell(row=r, column=6).number_format = "0.0%"
    wr.conditional_formatting.add("F5:F15", FormulaRule(formula=['AND(ISNUMBER(F5),ABS(F5)>0.25)'], fill=PatternFill("solid", bgColor="FFE08A", fgColor="FFE08A"), font=Font(bold=True, color="5A4300")))
    wr.freeze_panes = "B5"
    readme(wb, "Schedule III ratios", "Division II of Schedule III, Companies Act, 2013 (the ratios to be disclosed in the notes to the balance sheet)",
           ["1. Fill in the Inputs sheet from the financial statements. Use the same unit (for example Rs. lakh) in every cell.", "2. The Ratios sheet gives the eleven ratios for the current and the previous year, and the change.", "3. A change of more than 25% turns yellow. Schedule III needs the company to explain it in the notes. Type the reason in the last column.", "4. The formulas follow the numerators and denominators that Schedule III gives. Read the notes to Schedule III for the exact definitions of each item before you rely on the result."])
    save(wb, "schedule-iii-ratios.xlsx", "Schedule III ratios", "The eleven ratios a company must disclose, for two years, with the change and a place to give the reason.", "Division II, Schedule III, Companies Act, 2013", ["Inputs", "Ratios"])

    REG_ROWS = [
        ("Register of members", "Section 88(1)(a); Rule 3", "Form MGT-1", "register-of-members.xlsx"),
        ("Index of members", "Section 88(3)", "No form", "Sort the Register of members by name, and keep it as the index"),
        ("Register of debenture holders and other security holders", "Section 88(1)(b) and (c)", "Form MGT-2", "register-of-debenture-holders.xlsx"),
        ("Register of renewed and duplicate share certificates", "Section 46; Rule 6", "Form SH-2", "register-of-renewed-duplicate-certificates.xlsx"),
        ("Register of sweat equity shares", "Section 54; Rule 8", "Form SH-3", "register-of-sweat-equity-shares.xlsx"),
        ("Register of employee stock options", "Section 62(1)(b); Rule 12", "Form SH-6", "register-of-employee-stock-options.xlsx"),
        ("Register of shares or other securities bought back", "Section 68; Rule 17", "Form SH-10", "register-of-buy-back.xlsx"),
        ("Register of deposits", "Sections 73 to 76; Rule 14", "No form", "register-of-deposits.xlsx"),
        ("Register of charges", "Section 85; Rule 10", "Form CHG-7", "register-of-charges.xlsx"),
        ("Register of significant beneficial owners in the company", "Section 90; SBO Rules, 2018", "Form BEN-3", "register-of-significant-beneficial-owners.xlsx"),
        ("Register of directors and key managerial personnel and their shareholding", "Section 170; Rule 17", "No form", "register-of-directors-and-kmp.xlsx"),
        ("Register of directors' disclosures of interest", "Section 184", "Forms MBP-1, DIR-8, DIR-2", "register-of-director-disclosures.xlsx"),
        ("Register of investments not held in the company's name", "Section 187(4); Rule 14", "Form MBP-3", "register-of-investments-not-in-company-name.xlsx"),
        ("Register of loans, guarantee, security and acquisition made by the company", "Section 186(9); Rule 12", "Form MBP-2", "register-of-loans-guarantees-investments.xlsx"),
        ("Register of contracts or arrangements in which directors are interested", "Section 189; Rule 16", "Form MBP-4", "register-of-related-party-contracts.xlsx"),
        ("Register of allotments", "Sections 39 and 42", "Record, no form", "register-of-allotments.xlsx"),
        ("Register of transfers of shares", "Sections 56 and 58", "Record, no form", "share-transfer-register.xlsx"),
        ("Register of dividends", "Sections 123 to 125", "Record, no form", "register-of-dividends.xlsx"),
        ("Foreign register of members or debenture holders", "Section 88(4)", "Rule 4", "Not included: needed only if the company keeps a register outside India"),
        ("Minute books of Board, committee and general meetings", "Section 118; Secretarial Standards 1 and 2", "No form", "Not included: use the Minutes generator and keep the signed pages in a bound book"),
    ]
    reg_table = "<h2>Registers the Companies Act, 2013 requires</h2><p class=\"small\">Each register, the provision it comes from, its form, and the file in this library that covers it.</p><div style=\"overflow-x:auto\"><table class=\"tbl\"><thead><tr><th>Register</th><th>Provision</th><th>Form</th><th>File</th></tr></thead><tbody>"
    for nm, prov, frm, f in REG_ROWS:
        cell = f'<a href="{f}" download>Download</a>' if f.endswith(".xlsx") else f
        reg_table += f"<tr><td>{nm}</td><td>{prov}</td><td>{frm}</td><td>{cell}</td></tr>"
    reg_table += "</tbody></table></div>"
    BASIS = {f: prov + "; " + (frm if frm.startswith("Form") else "no prescribed form") for nm, prov, frm, f in REG_ROWS if f.endswith(".xlsx")}
    order = ["Statutory registers", "Filings, meetings and planners", "Checklists", "Accounts and finance", "Practice"]
    sections = ""
    for gname in order:
        items = sorted([m for m in made if m["group"] == gname], key=lambda m: 0 if "all-in-one" in m["file"] else 1)
        if items:
            sections += f"<h2>{gname}</h2><div class=\"xl\">" + "".join(
                f'<div class="card"><b>{m["name"]}</b><span>{m["blurb"]}</span><i>Based on: {BASIS.get(m["file"], m["basis"])}</i><a class="btn" href="{m["file"]}" download>Download .xlsx</a></div>' for m in items) + "</div>"
    body = f"""<h1>Excel utilities</h1><p class="lede">Free working files for the Company Secretary's desk: statutory registers, trackers, planners and checklists. Each file opens in Excel, LibreOffice and Google Sheets. Cells you fill in are yellow; cells that work themselves out are grey-blue.</p>
{sections}
<h2>How to use them</h2><ul><li>Download a file and save it with your company's name.</li><li>Fill in the yellow cells. The dates, totals and checks update on their own.</li><li>Check each rule and date against the current law. These files do not include extensions or exemptions announced by the Ministry.</li></ul>"""
    (out / "index.html").write_text(shell("Excel utilities | Company Law Knowledge Lab", "Statutory registers, compliance trackers and planners as free Excel files.", "xl", body), encoding="utf-8")
    print(f"{len(made)} Excel files written")


if __name__ == "__main__":
    for step in (build_generators, build_calendar, build_stamp, build_excel):
        try:
            step()
        except Exception as e:
            print(f"{step.__name__} ended with an error: {e}")
    sys.exit(0)
