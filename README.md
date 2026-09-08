# Apartment-Schedule
Apartment Schedule from AutoCAD block with attributes to excel using MVB

**[Template Row Manual](TEMPLATE-ROW.md)** — how the mapping rows on the `template` sheet decide which columns each output sheet gets, and the sigma / percent markers that control which of them are totalled.


Step 1:

![image](https://github.com/user-attachments/assets/3801c58d-e575-4391-8830-da2fbfa80abc)

Label the apratments in AutoCAD with the blocks provided. Fill out the apartment details in the attribute fields. For simplicity, it is best to do so in a separate drawing, with the main plan drawing xref'ed in. Putting diffenent types onto separate layers can ease management. Keeping a duplicate of each type assembled in one place can be also helpful, make sure to assign them to Block 0 to avoid them being counted

Step 2: 

Select all the tag blocks (selectsimilar) is a command that can help with it, and use Express Tools > Blocks > Export Attributes. Choose a location to save the export .txt to

Step 3:

Open the resulting .txt in Excel

![image](https://github.com/user-attachments/assets/6bf7ed16-7685-410e-97a7-d38852dd55f7)

Choose Delimited

![image](https://github.com/user-attachments/assets/c53ce7bb-da87-41c6-9c34-ec059d0b1330)

Use Tab as delimter

Click through the rest of the dialog and click Finish

Step 4:

Open the Template file. When first opening it, you might need to enable macros:
![image](https://github.com/user-attachments/assets/8aeaa461-ae9a-4b17-95c6-97a0aa73b2f5)


Step 5:

Copy all the data from text file into the sourceData sheet in the template file

Step 6:

Press alt + F8 and press run
![image](https://github.com/user-attachments/assets/87f2f436-6f11-408f-9b64-eb9897d9e562)


The scrpit converts the source spreadsheet looking like this:

![image](https://github.com/user-attachments/assets/69c3aa9f-5315-4ba4-8017-7a7ef6982824)

to a more user firendly format, broken down by floors and blocks:

## Changing the columns

The macros do not hard-code a column layout. Each output sheet reads one row of the `template` sheet — row 9 for the Long schedule, row 18 for the Short, row 29 for the unit types block — and builds itself from the field names it finds there. Reordering those cells reorders the schedule; deleting one drops the column; adding a sigma to the end of a cell totals it on every summary row.

See **[TEMPLATE-ROW.md](TEMPLATE-ROW.md)** for the full rules. A condensed version can be stamped into the workbook itself as cell notes on the mapping rows — run `WriteTemplateMemo` from `Template Memo.bas`.
