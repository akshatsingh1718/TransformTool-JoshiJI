# Work For 18 Feb

1. [x] if amount is 0 then dont include the row
2. [x] use strip in Party Name
3. [x] Add bill no for sale jsp/22/23/<starting no>
4. [x] Calculate CGST/SGST if gst no starts with "05.." else calculate IGST.

# Work for 19 Feb

1. [x] Product's Name = =concat("Medicine ", GSTPER, "%") | {Sale, Purchase} transform
2. [x] {CGST,SGST}AMT = round(2) | {Sale, Purchase} transform
3. [x] reg Type = "regular" if gst else "unregistered consumer" on all rows | {Sale, Purchase} transform 
4. [x] State Column =  {Purchase} transform | Link:https://www.zoho.com/in/books/kb/gst/valid-state-codes-list.html
    - Get the first 2 char of gst and lookup in valid-states-code-list and add the state name to state column
    - if unregistered then state="UK"
5. [x] 3004 to hsn| {Sale, Purchase} transform3004
6. [x] remove 5 char from party name. Purchase
7. [x] Sale party name 2x times appearings.
8. [x] In sale, remove double name using "-" 


# Work for 1 Mar 24
- [x] In GST: changed the Qty to use ceil values.
- [x] In GST: converted gst_percentage to int.
- [x] Implemented jj/only.
- [x] implemented IPD & OPD.
- [x] Chaned name of Sale app= GSTR1; Purchase App = GSTR2.
- [x] sale transformation name = GST Sale without qty 
- [x] gst transformation name = GSTR1 with qty
- [x] GSTR1 with qty transformation : Party = Swip card if inv no in mapping file

# Work for 2 Mar 24
- [x] Narration column in IPD and OPD.
- [x] implemented GST 2B Excels
- [x] Check missing row; in stock; change name to mutual fund. (Solved: stock.split("-")[:-1] is giving blank for non "-" texts )
- [x] ECHS DUE.
- [x] IPD/OPD remove blank rows


# Work for 7 Mar 24
- [x] GST 2B Excel: "Trade/Legal name" sort 
- [x] GSTR1 Self: check email for structure.