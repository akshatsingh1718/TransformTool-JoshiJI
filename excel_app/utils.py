import pandas as pd
from openpyxl.utils import get_column_letter
import copy
from datetime import datetime
import os
import json
import math
from typing import List

import warnings

warnings.simplefilter("ignore")


class BaseTransformExcel:

    @staticmethod
    def read_config(path: str):
        if os.path.exists(path):
            with open(path, "r") as f:
                return json.load(f)
        return None

    @staticmethod
    def get_filename_with_datetime(prefix="", suffix="", ext="xlsx"):
        # Get the current date and time
        current_datetime = datetime.now()
        # Format the date and time as desired
        formatted_datetime = current_datetime.strftime("%Y-%m-%d_%H-%M-%S")
        return f"{prefix}{formatted_datetime}{suffix}.{ext}"

    def get_default_row_format(cls, presets: dict) -> dict:
        row_dict: dict = cls.default_output_row

        for col in cls.output_columns:

            if col in presets.keys():
                row_dict[col] = presets[col]
                continue

            if col not in row_dict.keys():
                row_dict[col] = ""

        return row_dict

    def get(cls, row, column: str):
        return row[cls.column_to_idx[column]]

    @staticmethod
    def strip_column_names(df):
        """
        Strips leading and trailing whitespaces from column names in a DataFrame.

        Parameters:
        df (DataFrame): The pandas DataFrame.

        Returns:
        DataFrame: The DataFrame with stripped column names.
        """
        df.columns = [col.strip() for col in df.columns]
        return df


class EchsDueTransfromation(BaseTransformExcel):
    NAME = "ECHS DUE"
    CONFIG = "echs_due-config.json"

    def __init__(cls, config):
        cls.file_save_dir = config.get("file_save_dir")
        cls.output_columns = config.get("output_columns")
        cls.sheetname = config.get("sheetname")
        cls.file_prefix = config.get("file_prefix")
        cls.save_sheet_name = config.get("save_sheet_name")

    def transform(cls, file_paths: List, save=True) -> pd.DataFrame:
        # Read each file, skipping the first 7 rows
        read_df = lambda path: pd.read_excel(
            path, sheet_name=cls.sheetname, skiprows=2, skipfooter=1, header=None
        )

        df_list = []
        for path in file_paths:
            df = read_df(path)
            df["Date"] = "".join(os.path.basename(str(path)).split(".")[:-1])
            df_list.append(df)

        df = pd.concat(df_list, axis=0)
        df.columns = cls.output_columns

        filename = cls.get_filename_with_datetime(prefix=cls.file_prefix)
        xl_save_path = os.path.join(cls.file_save_dir, filename)

        # if dir not exist then create one
        if not os.path.isdir(cls.file_save_dir):
            os.makedirs(cls.file_save_dir, exist_ok=True)
        if save:
            print(f"File saved to: {xl_save_path}")
            df.to_excel(xl_save_path, sheet_name=cls.save_sheet_name, index=False)

        return dict(
            xl_save_path=xl_save_path, save_dir=cls.file_save_dir, xl_file_name=filename
        )


class GST2BTransfromation(BaseTransformExcel):
    NAME = "GST 2B Excel"
    CONFIG = "gst2bexcel_config.json"

    def __init__(cls, config):
        cls.file_save_dir = config.get("file_save_dir")
        cls.output_columns = config.get("output_columns")
        cls.sheetname = config.get("sheetname")
        cls.file_prefix = config.get("file_prefix")
        cls.save_sheet_name = config.get("save_sheet_name")

    def transform(cls, file_paths: List, save=True) -> pd.DataFrame:
        # Read each file, skipping the first 7 rows
        read_df = lambda path: pd.read_excel(path, sheet_name=cls.sheetname, skiprows=7)

        df_list = []
        for path in file_paths:
            for _, row in read_df(path).iterrows():
                df_list.append(list(row))

        df = pd.DataFrame(data=df_list, columns=cls.output_columns)

        filename = cls.get_filename_with_datetime(prefix=cls.file_prefix)
        xl_save_path = os.path.join(cls.file_save_dir, filename)

        # if dir not exist then create one
        if not os.path.isdir(cls.file_save_dir):
            os.makedirs(cls.file_save_dir, exist_ok=True)
        if save:
            # df.to_excel(xl_save_path, sheet_name=cls.save_sheet_name, index=False)
            print(f"File saved to: {xl_save_path}")
            df.to_excel(xl_save_path, sheet_name="Sheet1", index=False)

        return dict(
            xl_save_path=xl_save_path, save_dir=cls.file_save_dir, xl_file_name=filename
        )


class IpdOpdTransfromation(BaseTransformExcel):

    def __init__(cls, columns: str, _for: str, filename_prefix=""):
        cls.columns = columns
        cls.filename_prefix = filename_prefix
        cls.file_save_dir = "transformed"
        cls._for = _for

    def transform(cls, path: str, save=True) -> pd.DataFrame:
        df = pd.read_excel(path)

        df = cls.strip_column_names(df)

        df = df.sort_values(by=cls.columns)

        ## change date time fileds format
        # List to store the names of datetime columns
        datetime_columns = []
        # Identify datetime columns
        for column in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[column]):
                datetime_columns.append(column)
        # Convert datetime columns to date format="%d/%m/%Y"
        for column in datetime_columns:
            df[column] = df[column].dt.strftime("%d/%m/%Y")

        def concat_cols(r):
            if cls._for == "ipd":
                try:
                    return (
                        f"Bill No: {int(r['Bill No'])} {r['Patient Name']} {r['TPA']}"
                    )
                except Exception as e:
                    return ""
            try:
                return f"Bill No: {int(r['Bill No'])} {r['Patient Name']}"
            except Exception as e:
                return ""

        # concat columns
        df["Narration"] = df.apply(concat_cols, axis=1)

        # drop all the rows where all the cell are empty
        df.dropna(how="all", inplace=True)

        filename = cls.get_filename_with_datetime(prefix=f"{cls.filename_prefix}_")
        xl_save_path = os.path.join(cls.file_save_dir, filename)

        # if dir not exist then create one
        if not os.path.isdir(cls.file_save_dir):
            os.makedirs(cls.file_save_dir, exist_ok=True)

        if save:
            print(f"File saved to: {xl_save_path}")
            # Convert the dataframe to an XlsxWriter Excel object.
            df.to_excel(xl_save_path, sheet_name="Sheet1", index=False)
        return dict(
            xl_save_path=xl_save_path, save_dir=cls.file_save_dir, xl_file_name=filename
        )


class TransformExcelPurchase(BaseTransformExcel):

    def __init__(cls, config: dict) -> None:
        cls.output_columns = config["output_columns"]
        cls.gst_data = config["gst_data"]
        cls.default_output_row = config["default_output_row"]
        cls.columns_to_sum = config["columns_to_sum"]
        cls.target_columns_index = config["target_columns_index"]

        # constants
        cls.file_save_dir = "transformed"
        cls.filename_prefix = "GSTR2"

        cls.perfix_for_totals = config["perfix_for_totals"]
        cls.party_gst_index = config["party_gst_index"]
        cls.gst_mapping_xl_path: str = config["gst_mapping_xl_path"]
        cls.column_to_idx: dict = config["column_to_idx"]

    def get_gst_state(cls, gst_code: str) -> str:
        try:
            # Get the absolute path to the Excel file
            excel_file_path = os.path.join(cls.gst_mapping_xl_path)

            # Read the Excel file into a DataFrame
            df = pd.read_excel(excel_file_path)

            # Check if the input_value exists in the 'Mapping from' column
            mapping_to = df.loc[df["TIN"] == gst_code, "State"].values

            if len(mapping_to) > 0:
                # Return the corresponding mapping to value
                return mapping_to[0]
            else:
                # If input_value not found, return None
                return None
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def get_transformed_rows(cls, row):
        # if start with 05 then calculate cgst/sgst
        # if not start with 05 then calculate igst
        gst_no = str(row.iloc[cls.party_gst_index]).strip()
        calculate_igst = not (str(gst_no).strip().startswith("05"))
        state = "" if gst_no == "nan" else cls.get_gst_state(int(gst_no[:2]))

        new_rows = []
        # for per, amount in [(5, amount_5), (12, amount_12), (18, amount_18)]:
        for gst in cls.gst_data:
            gst_percentage = gst["gst_percentage"]
            cgst_sgst_per = (gst_percentage / 2) if (not calculate_igst) else 0.0
            amount = float(row.iloc[gst["target_column_idx"]])

            if amount == 0.0:
                continue

            cgst_sgst_amt = (
                round(amount * cgst_sgst_per / 100, 2) if not calculate_igst else 0.0
            )
            igst_amt = (
                round(amount * gst_percentage / 100, 2) if calculate_igst else 0.0
            )
            discount = 0.0
            total_amt = round(amount + (2 * cgst_sgst_amt) + igst_amt - discount, 2)
            new_row = cls.get_default_row_format(
                {
                    **{
                        i["column"]: str(row.iloc[i["target_column_idx"]]).strip()
                        for i in cls.target_columns_index
                    },
                    "Party's GST": "" if gst_no == "nan" else gst_no,
                    "Reg Type": (
                        "unregistered/consumer" if gst_no == "nan" else "regular"
                    ),
                    "Product's Name": f"Medicine {gst_percentage}%",
                    "State": "Uttrakhand" if gst_no == "nan" else state,
                    "Party/Cash": str(row.iloc[cls.column_to_idx["Party/Cash"]])[
                        5:
                    ].strip(),
                    "GST%": gst_percentage,
                    "Amount": amount,
                    "CGST %": cgst_sgst_per,
                    "CGST Amt": cgst_sgst_amt,
                    "SGST %": cgst_sgst_per,
                    "SGST Amt": cgst_sgst_amt,
                    "IGST Amt": igst_amt,
                    "Discount %": discount,
                    "Total": total_amt,
                }
            )

            new_rows.append(copy.deepcopy(new_row))

        return new_rows

    def transform(cls, path: str, save=True):
        df = pd.read_excel(path)

        rows_to_df = []
        # Iterate over rows
        for _, row in df.iterrows():
            if not pd.notna(row.iloc[0]):
                continue

            # Check for date in 1st column
            try:
                pd.to_datetime(row.iloc[0], format="%d/%m/%Y").strftime(
                    format="%d/%m/%Y"
                )
            except Exception as e:
                continue

            # If the row starts with a bill number
            rows_to_df += cls.get_transformed_rows(row)

        filename = cls.get_filename_with_datetime(prefix=f"{cls.filename_prefix}_")
        xl_save_path = os.path.join(cls.file_save_dir, filename)

        # if dir not exist then create one
        if not os.path.isdir(cls.file_save_dir):
            os.makedirs(cls.file_save_dir, exist_ok=True)

        writer = pd.ExcelWriter(xl_save_path, engine="xlsxwriter")

        result_df = pd.DataFrame(rows_to_df, columns=cls.output_columns)

        # Convert the dataframe to an XlsxWriter Excel object.
        result_df.to_excel(writer, sheet_name="Sheet1", index=False)

        # Get the xlsxwriter objects from the dataframe writer object.
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        no_of_row = worksheet.dim_rowmax
        for col_to_sum in cls.columns_to_sum:
            col_letter = get_column_letter(
                result_df.columns.get_loc(col_to_sum) + 1
            )  # offset of 1 for index to pos

            index_of_sheet = f"{col_letter}{no_of_row + 2}"
            formula = f"=SUM({col_letter}2:{col_letter}{no_of_row + 1})"

            worksheet.write_formula(index_of_sheet, formula)

        if len(cls.columns_to_sum) > 0:
            col_letter = get_column_letter(
                result_df.columns.get_loc(cls.perfix_for_totals["column"]) + 1
            )  # offset of 1 for index to pos
            index_of_sheet = f"{col_letter}{no_of_row + 2}"
            worksheet.write_string(index_of_sheet, cls.perfix_for_totals["label"])

        if save:
            print(f"File saved to: {xl_save_path}")
            writer.close()
        return dict(
            xl_save_path=xl_save_path, save_dir=cls.file_save_dir, xl_file_name=filename
        )


class TransformExcelSale(BaseTransformExcel):

    def __init__(cls, config: dict, *args, **kwargs) -> None:
        cls.output_columns = config["output_columns"]
        cls.gst_data = config["gst_data"]
        cls.default_output_row = config["default_output_row"]
        cls.columns_to_sum = config["columns_to_sum"]
        cls.target_columns_index = config["target_columns_index"]
        cls.file_save_dir = "transformed"
        cls.filename_prefix = "GSTR1"
        cls.perfix_for_totals = config["perfix_for_totals"]
        cls.party_gst_index = config["party_gst_index"]
        cls.column_to_idx: dict = config["column_to_idx"]

        # changes every time
        cls.bill_no_prefix = kwargs.get("bill_no_prefix", None)
        cls.bill_no_suffix_counter = kwargs.get("bill_no_suffix_counter", None)
        cls.calculate_igst = kwargs.get("calculate_igst", False)

    def get_transformed_rows(
        cls, row, bill_date: str, bill_no: str, calculate_igst=False
    ):
        new_rows = []
        gst_no = str(row.iloc[cls.party_gst_index])

        for gst in cls.gst_data:
            gst_percentage = gst["gst_percentage"]
            cgst_sgst_per = gst_percentage / 2
            amount = float(row.iloc[gst["target_column_idx"]])
            if amount == 0.0:
                continue

            cgst_sgst_amt = round(amount * cgst_sgst_per / 100, 2)
            igst_amt = (
                round(amount * gst_percentage / 100, 2) if calculate_igst else 0.0
            )
            discount = 0.0
            total_amt = round(amount + (2 * cgst_sgst_amt) + igst_amt - discount, 2)
            new_row = cls.get_default_row_format(
                {
                    **{
                        i["column"]: str(row.iloc[i["target_column_idx"]]).strip()
                        for i in cls.target_columns_index
                    },
                    "Reg Type": (
                        "unregistered/consumer" if gst_no == "nan" else "regular"
                    ),
                    "Product's Name": f"Medicine {gst_percentage}%",
                    "Party/Cash": str(row.iloc[cls.column_to_idx["Party/Cash"]])
                    .split("-")[0]
                    .strip(),
                    "Bill No.": bill_no,
                    "Inv Date": bill_date,
                    "GST%": gst_percentage,
                    "Amount": amount,
                    "CGST %": cgst_sgst_per,
                    "CGST Amt": cgst_sgst_amt,
                    "SGST %": cgst_sgst_per,
                    "SGST Amt": cgst_sgst_amt,
                    "IGST Amt": igst_amt,
                    "Discount %": discount,
                    "Total": total_amt,
                }
            )
            new_rows.append(copy.deepcopy(new_row))

        return new_rows

    def transform(cls, path: str, save=True):
        df = pd.read_excel(path)
        rows_to_df = []
        bill_counter = cls.bill_no_suffix_counter

        # Iterate over rows
        for _, row in df.iterrows():
            # If the row starts with a date (assuming the date is in the first cell)
            if pd.notna(row.iloc[0]):
                try:
                    current_bill_date = pd.to_datetime(
                        row.iloc[0], format="%d/%m/%Y"
                    ).strftime(format="%d/%m/%Y")
                    continue
                except ValueError:
                    # If the value is not a valid date, move to the next row
                    pass

            # If the row starts with a bill number
            if (
                str(row.iloc[0]).startswith("A")
                and len(str(row.iloc[0]).split(" ")) == 1
                and len(str(row.iloc[0])) >= 6
            ):
                # Set current bill number
                new_tranformed_rows = cls.get_transformed_rows(
                    row,
                    bill_date=current_bill_date,
                    bill_no=f"{cls.bill_no_prefix}{bill_counter:05d}",
                    calculate_igst=cls.calculate_igst,
                )
                rows_to_df += new_tranformed_rows

                # if new rows added then only increase the bill counter
                # there may be case where no rows were added because amount is 0
                if len(new_tranformed_rows) > 0:
                    bill_counter += 1

        filename = cls.get_filename_with_datetime(prefix=f"{cls.filename_prefix}_")
        xl_save_path = os.path.join(cls.file_save_dir, filename)

        # if dir not exist then create one
        if not os.path.isdir(cls.file_save_dir):
            os.makedirs(cls.file_save_dir, exist_ok=True)

        writer = pd.ExcelWriter(xl_save_path, engine="xlsxwriter")

        result_df = pd.DataFrame(rows_to_df, columns=cls.output_columns)

        # Convert the dataframe to an XlsxWriter Excel object.
        result_df.to_excel(writer, sheet_name="Sheet1", index=False)

        # Get the xlsxwriter objects from the dataframe writer object.
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        no_of_row = worksheet.dim_rowmax
        for col_to_sum in cls.columns_to_sum:
            col_letter = get_column_letter(
                result_df.columns.get_loc(col_to_sum) + 1
            )  # offset of 1 for index to pos

            index_of_sheet = f"{col_letter}{no_of_row + 2}"
            formula = f"=SUM({col_letter}2:{col_letter}{no_of_row + 1})"

            worksheet.write_formula(index_of_sheet, formula)

        if len(cls.columns_to_sum) > 0:
            col_letter = get_column_letter(
                result_df.columns.get_loc(cls.perfix_for_totals["column"]) + 1
            )  # offset of 1 for index to pos
            index_of_sheet = f"{col_letter}{no_of_row + 2}"
            worksheet.write_string(index_of_sheet, cls.perfix_for_totals["label"])

        if save:
            print(f"File saved to: {xl_save_path}")
            writer.close()

        return dict(
            xl_save_path=xl_save_path, save_dir=cls.file_save_dir, xl_file_name=filename
        )

# Stock Transformation earlier
class TransformStockExcel(BaseTransformExcel):
    APP_NAME = "Mutual Fund"

    def __init__(cls, config: dict, *args, **kwargs) -> None:
        cls.file_save_dir = "transformed"

        pass

    def transform(cls, path: str, save=True) -> pd.DataFrame:
        """
        - Iterate over df rows
        - If row == header then append opening balalnce data.
        """
        df = pd.read_excel(path, skiprows=4)
        result = []
        stock_name = None
        for index, row in df.iterrows():
            # if the header found
            if (
                row[0] == "Sr."
                and row[1] == "Transaction Date"
                and row[2] == "Transaction Type"
                and row[3] == "Amount"
                and row[4] == "Units"
            ):
                # if header found then get stock name from prv cell
                stock_name = str(df.iloc[index - 1, 0]).strip().split("-")
                stock_name = "".join(stock_name[:-1]) if len(stock_name) > 1 else stock_name[0]
                # if stock_name == "":
                #     [print("==============================")]
                #     print(f"Stock found at {index - 1}")
                #     print(str(df.iloc[index - 1, 0]).split("-")[:-1])
                # also append the next cell data to the result which is opening balance
                result.append(
                    [
                        len(result) + 1,
                        stock_name,
                        "",
                        df.iloc[index + 1, 1],
                        df.iloc[index + 1, 3],
                        "",
                    ]
                )
            # if the cell is not null and has int data at first
            elif (
                pd.notnull(row[0]) and isinstance(row[0], int) and str(row[2]) != "nan"
            ):

                result.append(
                    [len(result) + 1, stock_name, row[1], row[2], row[3], row[4]]
                )
                # if stock_name == "":
                #     print("find")
            # if nothing matches then continue
            else:
                continue

        columns = [
            "Sr.",
            "Fund Name",
            "Transaction Date",
            "Transaction Type",
            "Amount",
            "Units",
        ]

        df = pd.DataFrame(result, columns=columns)

        if save:
            filename = cls.get_filename_with_datetime(prefix="Mutual-Fund_")
            xl_save_path = os.path.join(cls.file_save_dir, filename)
            df.to_excel(xl_save_path, index=False)
            print(f"File saved to: {xl_save_path}")

        return dict(
            xl_save_path=xl_save_path, save_dir=cls.file_save_dir, xl_file_name=filename
        )


class TransformExcelGST(BaseTransformExcel):
    NAME = "GSTR1 With Qty"
    CONFIG = "gst_config.json"

    def __init__(cls, config: dict, *args, **kwargs) -> None:
        cls.file_save_dir = "transformed"
        cls.output_columns = config["output_columns"]
        cls.default_output_row = config["default_output_row"]
        cls.columns_to_sum = config["columns_to_sum"]
        cls.perfix_for_totals = config["perfix_for_totals"]
        cls.column_to_idx: dict = config["column_to_idx"]
        cls.target_columns_index: list = config["target_columns_index"]

        # changes every time
        cls.bill_no_prefix = kwargs.get("bill_no_prefix", None)
        cls.bill_no_suffix_counter = kwargs.get("bill_no_suffix_counter", None)
        cls.calculate_igst = kwargs.get("calculate_igst", False)

        # Mapping file
        cls._mapping_df = kwargs.get("mapping_df", None)
        cls.mapping_sheetname = config["mapping_sheetname"]

    def transform_mapping(cls):

        df = cls._mapping_df.iloc[6:, 1:]
        # Set the headers to be the values from the first row
        new_headers = df.iloc[0]
        df.columns = new_headers

        # Reset the index
        df.reset_index(drop=True, inplace=True)

        cls._mapping_df = df

    def is_present_in_mapping(cls, value_to_find, column: str = "Bill No."):
        if value_to_find in cls._mapping_df[column].values:
            index = cls._mapping_df[cls._mapping_df[column] == value_to_find].index[0]
            if float(cls._mapping_df.at[index, "INDIAN BANK"]) > 0.0:
                return True
        return False

    def get_transformed_rows(cls, row, bill_no: str):
        new_rows = []

        cgst_sgst_per = float(row[cls.column_to_idx["SGST %"]])
        igst_per = float(row[cls.column_to_idx["IGST %"]])
        gst_percentage = int(cgst_sgst_per * 2)
        amount = float(row.iloc[cls.column_to_idx["Amount"]])

        if amount == 0.0:
            return []

        cgst_sgst_amt = round(amount * cgst_sgst_per / 100, 2)
        igst_amt = round(amount * igst_per / 100, 2)
        discount = 0.0
        total_amt = round(amount + (2 * cgst_sgst_amt) + igst_amt - discount, 2)
        party_cash = (
            "SWIP CARD"
            if cls.is_present_in_mapping(value_to_find=cls.get(row, "Inv No."))
            else cls.get(row, "Party/Cash")
        )
        new_row = cls.get_default_row_format(
            {
                **{
                    tg: str(row.iloc[cls.column_to_idx[tg]]).strip()
                    for tg in cls.target_columns_index
                },
                "Bill No.": bill_no,
                "Party/Cash": party_cash,
                "Qty": math.ceil(row[cls.column_to_idx["Qty"]]),
                "Reg Type": "unregistered/consumer",
                "Product's Name": f"Medicine {gst_percentage}%",
                "GST%": gst_percentage,
                "Amount": amount,
                "CGST %": cgst_sgst_per,
                "CGST Amt": cgst_sgst_amt,
                "SGST %": cgst_sgst_per,
                "SGST Amt": cgst_sgst_amt,
                "IGST Amt": igst_amt,
                "Discount %": discount,
                "Total": total_amt,
            }
        )
        new_rows.append(copy.deepcopy(new_row))

        return new_rows

    def transform(cls, path: str, save=True):
        df = pd.read_excel(path)
        # drop the 1st empty column
        # drop the 1st empty column
        df = df.iloc[:, 1:]

        cls.transform_mapping()

        rows_to_df = []
        bill_counter = cls.bill_no_suffix_counter - 1
        last_row = None

        # Iterate over rows
        for _, row in df.iterrows():
            # if gross total reached then break the loop

            if str(row.iloc[1]).strip() == "Gross Total":
                break

            # if cell-0 is nan and we have last row data
            if str(row.iloc[0]).strip() == "nan" and last_row is not None:
                for copy_idx in [1, 2, 3, 4, 5, 6, 7]:
                    row.iloc[copy_idx] = last_row.iloc[copy_idx]

            # if the row does not match the following format= "<digit><dot>"
            elif not str(row.iloc[0]).split(".")[0].strip().isdigit():
                continue
            else:
                # digit case found
                bill_counter += 1
                last_row = copy.deepcopy(row)

            new_tranformed_rows = cls.get_transformed_rows(
                row,
                bill_no=f"{cls.bill_no_prefix}{bill_counter:05d}",
            )

            # if new rows added then only increase the bill counter
            # there may be case where no rows were added because amount is 0
            if len(new_tranformed_rows) > 0:
                rows_to_df += new_tranformed_rows

        filename = cls.get_filename_with_datetime(prefix="GSTR1_W_Qty_")
        xl_save_path = os.path.join(cls.file_save_dir, filename)

        # if dir not exist then create one
        if not os.path.isdir(cls.file_save_dir):
            os.makedirs(cls.file_save_dir, exist_ok=True)

        writer = pd.ExcelWriter(xl_save_path, engine="xlsxwriter")

        result_df = pd.DataFrame(rows_to_df, columns=cls.output_columns)

        # Convert the dataframe to an XlsxWriter Excel object.
        result_df.to_excel(writer, sheet_name="Sheet1", index=False)

        # Get the xlsxwriter objects from the dataframe writer object.
        worksheet = writer.sheets["Sheet1"]

        no_of_row = worksheet.dim_rowmax
        for col_to_sum in cls.columns_to_sum:
            col_letter = get_column_letter(
                result_df.columns.get_loc(col_to_sum) + 1
            )  # offset of 1 for index to pos

            index_of_sheet = f"{col_letter}{no_of_row + 2}"
            formula = f"=SUM({col_letter}2:{col_letter}{no_of_row + 1})"

            worksheet.write_formula(index_of_sheet, formula)

        if len(cls.columns_to_sum) > 0:
            col_letter = get_column_letter(
                result_df.columns.get_loc(cls.perfix_for_totals["column"]) + 1
            )  # offset of 1 for index to pos
            index_of_sheet = f"{col_letter}{no_of_row + 2}"
            worksheet.write_string(index_of_sheet, cls.perfix_for_totals["label"])

        if save:
            print(f"File saved to: {xl_save_path}")
            writer.close()

        return dict(
            xl_save_path=xl_save_path, save_dir=cls.file_save_dir, xl_file_name=filename
        )


class TransformExcelJJOnly(BaseTransformExcel):
    NAME = "JJ/Only"

    def __init__(cls, config: dict, *args, **kwargs) -> None:
        # Product's Name: "Article Description" + " " + "EAN Number"
        # Amount : Rate * Qty
        # Invoice No: <prefix><counter>

        cls.output_columns = config["output_columns"]
        cls.default_output_row = config["default_output_row"]
        cls.columns_to_sum = config["columns_to_sum"]
        cls.target_columns_index = config["target_columns_index"]
        cls.file_save_dir = "transformed"
        cls.perfix_for_totals = config["perfix_for_totals"]
        cls.column_to_idx: dict = config["column_to_idx"]

        # changes every time
        cls.bill_no_prefix = kwargs.get("bill_no_prefix", None)
        cls.bill_no_suffix_counter = kwargs.get("bill_no_suffix_counter", None)

    def get_transformed_rows(cls, row, inv_no: str, calculate_igst=False):
        new_rows = []
        gst_percentage = int(cls.get(row, "GST %"))
        rate = round(float(cls.get(row, "Rate")), 2)
        qty = math.ceil(cls.get(row, "Qty"))
        amount = round(rate * qty, 2)
        discount = 0.0

        cgst_sgst_per = gst_percentage / 2

        if amount == 0.0:
            return None

        cgst_sgst_amt = round(amount * cgst_sgst_per / 100, 2)
        igst_amt = round(amount * gst_percentage / 100, 2) if calculate_igst else 0.0
        total_amt = round(amount + (2 * cgst_sgst_amt) + igst_amt - discount, 2)
        inv_date = cls.get(row, "Inv Date").strftime(format="%d/%m/%Y")

        # party/cash
        party_cash = (
            cls.get(row, "Article Description").strip()
            + " "
            + cls.get(row, "EAN Number").strip()
        )

        new_row = cls.get_default_row_format(
            {
                **{
                    colname: str(cls.get(row, colname)).strip()
                    for colname in cls.target_columns_index
                },
                "Inv Date": inv_date,
                "Inv No.": inv_no,
                "Product's Name": party_cash,
                "GST %": gst_percentage,
                "Amount": amount,
                "CGST %": cgst_sgst_per,
                "CGST Amt": cgst_sgst_amt,
                "SGST %": cgst_sgst_per,
                "SGST Amt": cgst_sgst_amt,
                "IGST Amt": igst_amt,
                "Discount %": discount,
                "Total": total_amt,
            }
        )
        new_rows.append(copy.deepcopy(new_row))

        return new_rows

    def transform(cls, path: str, save=True):
        df = pd.read_excel(path)
        rows_to_df = []
        bill_counter = cls.bill_no_suffix_counter

        # Removing the first 2 rows and first col
        df = df.iloc[2:, 1:]
        df.reset_index(drop=True, inplace=True)

        billNo_counter_dict = dict()

        # Iterate over rows
        for _, row in df.iterrows():
            row = list(row)

            if billNo_counter_dict.get(cls.get(row, "Bill No.")) is not None:
                _bill_counter = billNo_counter_dict.get(cls.get(row, "Bill No."))
            else:
                _bill_counter = bill_counter
                # add the new bill number to dict
                billNo_counter_dict[cls.get(row, "Bill No.")] = _bill_counter
                bill_counter += 1  # update the bill no

            # Set current bill number
            new_tranformed_rows = cls.get_transformed_rows(
                row,
                inv_no=f"{cls.bill_no_prefix}{_bill_counter:05d}",
            )
            rows_to_df += new_tranformed_rows

        # File saving starts from here
        filename = cls.get_filename_with_datetime(prefix="JJOnly_")
        xl_save_path = os.path.join(cls.file_save_dir, filename)

        # if dir not exist then create one
        if not os.path.isdir(cls.file_save_dir):
            os.makedirs(cls.file_save_dir, exist_ok=True)

        writer = pd.ExcelWriter(xl_save_path, engine="xlsxwriter")

        result_df = pd.DataFrame(rows_to_df, columns=cls.output_columns)

        # Convert the dataframe to an XlsxWriter Excel object.
        result_df.to_excel(writer, sheet_name="Sheet1", index=False)

        # Get the xlsxwriter objects from the dataframe writer object.
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        no_of_row = worksheet.dim_rowmax
        for col_to_sum in cls.columns_to_sum:
            col_letter = get_column_letter(
                result_df.columns.get_loc(col_to_sum) + 1
            )  # offset of 1 for index to pos

            index_of_sheet = f"{col_letter}{no_of_row + 2}"
            formula = f"=SUM({col_letter}2:{col_letter}{no_of_row + 1})"

            worksheet.write_formula(index_of_sheet, formula)

        if len(cls.columns_to_sum) > 0:
            col_letter = get_column_letter(
                result_df.columns.get_loc(cls.perfix_for_totals["column"]) + 1
            )  # offset of 1 for index to pos
            index_of_sheet = f"{col_letter}{no_of_row + 2}"
            worksheet.write_string(index_of_sheet, cls.perfix_for_totals["label"])

        if save:
            print(f"File saved to: {xl_save_path}")
            writer.close()

        return dict(
            xl_save_path=xl_save_path, save_dir=cls.file_save_dir, xl_file_name=filename
        )


## Sanity check functions


def check_for_stock():
    default_config = {}
    obj = TransformStockExcel(default_config)
    obj.transform("/home/akshat/Documents/projects/joshi-uncle/data/MutalFund_stock_forErrorRows.xlsx", save=True)


def check_for_purchase():
    default_config = TransformExcelPurchase.default_config()

    obj = TransformExcelPurchase(default_config)

    def get_df():
        sale_path = "/home/akshat/Documents/projects/joshi-uncle/data/REPORT_2222.XLS"
        df = pd.read_excel(
            sale_path, sheet_name="MARG ERP 9+ Excel Report", header=None
        )
        return df

    obj.transform_bill(get_df())


def check_for_sale():
    default_config = TransformExcelSale.default_config()

    obj = TransformExcelSale(default_config)

    def get_df():
        sale_path = "data/SALE JAN FINAL 2024.XLS"
        df = pd.read_excel(
            sale_path, sheet_name="MARG ERP 9+ Excel Report", header=None
        )
        return df

    obj.transform_bill(get_df())


def check_for_gst():
    default_config = TransformExcelGST.read_config(
        os.path.join("excel_app", "config", "gst_config.json")
    )
    bill_no_prefix = "JPS/23/24/"
    bill_no_suffix_counter = 1546

    obj = TransformExcelGST(
        default_config,
        bill_no_prefix=bill_no_prefix,
        bill_no_suffix_counter=bill_no_suffix_counter,
    )

    def get_df():
        sale_path = "../data/gst2.xls"
        df = pd.read_excel(sale_path, header=None)
        return df

    obj.transform(get_df())


def check_for_jjonly():

    default_config = TransformExcelSale.read_config(
        os.path.join("excel_app", "config", "jj-only_config.json")
    )
    bill_no_prefix = "TDR/23-24/"
    bill_no_suffix_counter = 1701

    obj = TransformExcelJJOnly(
        default_config,
        bill_no_prefix=bill_no_prefix,
        bill_no_suffix_counter=bill_no_suffix_counter,
    )

    def get_df():
        sale_path = "../data/jj-only.xlsx"
        df = pd.read_excel(sale_path, header=None)
        return df

    obj.transform(get_df())


def check_for_gst2b():
    default_config = TransformExcelSale.read_config(
        os.path.join("excel_app", "config", "gst2b_config.json")
    )
    obj = GST2BTransfromation(default_config)

    def get_df():
        dfs = []
        for file in os.listdir("../data/GST2BExcels/"):
            dfs.append("../data/GST2BExcels/" + file)

        return dfs

    dfs = get_df()
    obj.transform(dfs)


def check_for_echs():
    default_config = TransformExcelSale.read_config(
        os.path.join("excel_app", "config", "echs_due-config.json")
    )
    obj = EchsDueTransfromation(default_config)

    def get_df():
        dfs = []
        for file in os.listdir("../data/echs-due/"):
            dfs.append("../data/echs-due/" + file)

        return dfs

    dfs = get_df()
    obj.transform(dfs)


if __name__ == "__main__":
    check_for_stock()
