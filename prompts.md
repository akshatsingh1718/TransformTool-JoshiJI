Below the schema provided for the pandas dataframe.
```
Bill No.        : (str)
Inv No.         : (str)
Inv Date        : (str)
Party/Cash      : (str)
Product's Name  : (str)
HSN Code        : (str)
Qty             : (int)
Rate            : (int)
GST%            : (int)
Amount          : (float)
CGST %          : (float)
CGST Amt        : (float)
SGST %          : (float)
SGST Amt        : (float)
IGST Amt        : (float)
Discount %      : (float)
Total           : (float)
State           : (str)
```

Below is the psuedo code for the python pandas script:

1. filtered_df: Filter distinct values for "inv Date" column and iterate over distinct values and filter results to filter_df.
2. filtered_df2: Then on filtered_df filter distinct values for "Party/Cash" column and iterate over them as well.
3. If for all the values of column "Party/Cash" in {"ECHS", "PMJAY", "SWIP CARD"}  do the following:
    - In filter_df3 do the following: 
        - Sum the int and float values.
        - for all the str values select one string value as they can not be summed up.
        - filter on distinct "Product's Name" as well and for all the filtered rows.
4. Else if for all the other  "Party/Cash":
    - filter_df3: filter on distinct "Product's Name" as well and for all the filtered rows.
    - In filter_df3 do the following:
        - Sum the int and float values.
        - for all the str values select one string value as they can not be summed up.
        - Rename the value of "Party/Cash" = "CASH".
5. result_df: Now filtered_3 will have only one row and append this row to a new result_df.

If you have any doubts and unclear about anything ask me.