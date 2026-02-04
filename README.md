
```markdown

\# “Plus” quotation generator – Technical documentation (VBA / Excel)



This “plus” version extends the quotation generator with extra automation features and a row-hiding macro.\[file:3]\[file:4]  

The goal is to simplify reading and printing the quotation by showing only the lines that are actually used.



\---



\## 1. Additional functional goals



Compared to the 2010 version:



\- Improve the user experience on the quotation sheet.  

\- Automatically hide unused lines (empty quantity or empty control cell).  

\- Prepare a clean quotation for printing / export (no “empty lines”).\[file:3]\[file:4]  



\---



\## 2. Workbook structure



The base structure remains similar to the 2010 version: catalog + quotation input sheet.\[file:4]  



\### 2.1. QUOTATION sheet (example)



\- Typical columns:  

&#x20; - Article Reference Code  

&#x20; - Article Label  

&#x20; - Unit of Measure  

&#x20; - Quantity  

&#x20; - Unit Price excl. tax  

&#x20; - Total amount  

&#x20; - “Control / activation” column indicating whether the row should be displayed.\[file:4]  



\- Row 4 (e.g. cell F4) is used as an entry point for filtering: the macro selects from this cell.\[file:3]  



\---



\## 3. Row-hiding macro



\### 3.1. VBA code



```vba

Private Sub CommandButton1\_Click()



&#x20;   Range("F4").Select

&#x20;   For Each o In Selection

&#x20;       If o.Value = "" Then

&#x20;           o.EntireRow.Hidden = True

&#x20;       End If

&#x20;   Next



End Sub



