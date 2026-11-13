section,case,court,year,citation,held,link
149,Name v. Name,Supreme Court,2015,(2015) 1 SCC 1,Held this point | Held that point,
# Guide Notes: what to put in this folder

Each file feeds one heading under "Guide Notes" on the section page. The website reads them every time it is rebuilt. Open a file with the pencil icon, add a line at the bottom, and commit.

The first column is always `section`. Write the section number (for example `149`) or several, such as `"149, 152"`. A line is shown under every section you list.
If a text has a comma, put it inside double quotes.

| File | Columns | Shown under |
|---|---|---|
| analogous.csv | section, old_section, note | Analogous Law (Companies Act, 1956) |
| lodr.csv | section, regulation, note | SEBI (LODR) Regulations, 2015 |
| exemptions.csv | section, applies_to, note, source | Exemption to certain companies |
| circulars.csv | section, title, date, reference, link | Circulars and clarifications |
| cases.csv | section, case, court, year, citation, held, link | Judicial pronouncements, and the Table of cases page |

In `cases.csv`, separate several held points with a `|` sign. The same case listed against many sections appears once in the Table of cases.

Only add what you have checked. The tool never invents cases, circulars or exemptions.
