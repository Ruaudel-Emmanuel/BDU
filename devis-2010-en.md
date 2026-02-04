\# 2010 Quotation Generator – Technical Documentation (VBA / Excel)



This Excel workbook is a quotation generator based on a detailed unit price schedule for landscaping and civil engineering works (City of Rennes, etc.).\[file:2]  

It is one of the author’s first \*\*automation\*\* projects: the goal is to transform a complex price catalog into an interactive quotation-building tool.



\---



\## 1. Functional goals



\- Centralize the unit price schedule in a single workbook.  

\- Allow fast selection of items (codes, labels, units, unit prices).  

\- Automatically compute amounts based on entered quantities.  

\- Structure the quotation by families (A, B, C, D…: earthworks, planting, networks, etc.).\[file:2]  



\---



\## 2. Workbook structure



> Sheet names may vary, but the logical structure is as follows.



\### 2.1. Catalog / price list sheet



\- Main columns:  

&#x20; - `Code Catalogue` (e.g. JAE)  

&#x20; - `Article Reference Code` (e.g. A, A1, A10001…)  

&#x20; - `Article Type (1 or 2)`  

&#x20; - `Article Label`  

&#x20; - `Unit Code` (F, u, m², m³, ml, etc.)  

&#x20; - `VAT Code`  

&#x20; - `Unit Price excl. tax`  

&#x20; - `Article Description` (long CCTP-like text).\[file:2]  



\- Hierarchy:  

&#x20; - Level 2 = families / sub-families (A, A1, B, B1…).  

&#x20; - Level 1 = priced items (full code such as A10001, B10101…).\[file:2]  



\### 2.2. Quotation input sheet



\- Input area:  

&#x20; - Item code (or dropdown list)  

&#x20; - Label  

&#x20; - Unit  

&#x20; - Unit price excl. tax  

&#x20; - Quantity  

&#x20; - Total amount (Qty × UP).\[file:2]  



\- Totals:  

&#x20; - Subtotals by chapter (A, B, C…).  

&#x20; - Grand total excl. tax, VAT, incl. tax (if implemented).



\### 2.3. Auxiliary sheets (optional)



\- Parameters (VAT, miscellaneous rates).  

\- Print-friendly quotation layout.



\---



\## 3. Automation logic (no full code)



\### 3.1. Filling quotation lines



\- Lookup of the item by code in the catalog sheet (via VLOOKUP / INDEX+MATCH or VBA macro).  

\- Automatic retrieval: label, unit, unit price, optional long description.\[file:2]  



\### 3.2. Automated calculations



\- Amount formula: `=Quantity \* Unit\_Price`.  

\- Aggregation of amounts per family (conditional sum on code A, B, C…).\[file:2]  



\### 3.3. Chapter structure



\- Use of chapter codes (A, B, C…) to display section headers in the quotation.  

\- Ability to include / exclude chapters according to the project.



\---



\## 4. VBA macros (concepts)



Typical responsibilities:



\- Initialize a new quotation (clear previous quantities / amounts).  

\- Copy a selection of items from the catalog to the quotation sheet.  

\- Update amounts and totals.



\---



\## 5. “Automation project” perspective



\- At that time, the “RPA / automation” vocabulary was not used, but the need already was:  

&#x20; - reduce quotation preparation time,  

&#x20; - secure prices and labels (single source of truth),  

&#x20; - structure a recurring business process (tenders, framework contracts, etc.).\[file:2]  



\- The Excel + VBA solution is a forerunner of what would now be implemented as:  

&#x20; - a web app,  

&#x20; - a catalog API,  

&#x20; - or a no-code workflow.



\### 5.1. Estimated time savings



Before this quotation generator was implemented, preparing a full quotation from the unit price schedule typically took between 1 h 30 and 2 h (manual lookup of items, copy/paste, checks).  

With the automated workbook (item lookup, formulas and structure ready to use), the same quotation can be prepared in about 30 to 45 minutes, i.e. roughly 1 h saved per quotation.  



On a conservative basis of 40 to 60 quotations per year, this represents around 40 to 60 working hours saved annually, only on the quotation preparation phase.  

In a public-sector context, this is expressed as saved time and improved focus on technical analysis and stakeholder coordination rather than as direct financial savings.\[file:2]\[file:4]  



\---



\## 6. Modern evolution ideas



\- Export the quotation as a nicely formatted PDF.  

\- Connect to a database (SQLite, PostgreSQL, etc.) instead of Excel.  

\- Web interface (Django / FastAPI + front-end) reusing the same price list logic.  

\- Integration with automation tools (n8n, Make, Zapier) for full chain: request → pricing → PDF sending.



