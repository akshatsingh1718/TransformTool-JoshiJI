import pandas as pd
from openpyxl.utils import get_column_letter
import copy
from datetime import datetime
import os
import json

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


class TransformExcelPurchase(BaseTransformExcel):

    def __init__(cls, config: dict) -> None:
        cls.output_columns = config["output_columns"]
        cls.gst_data = config["gst_data"]
        cls.default_output_row = config["default_output_row"]
        cls.columns_to_sum = config["columns_to_sum"]
        cls.target_columns_index = config["target_columns_index"]
        cls.file_save_dir = "transformed"
        cls.perfix_for_totals = config["perfix_for_totals"]
        cls.party_gst_index = config["party_gst_index"]
        cls.gst_mapping_xl_path: str = config["gst_mapping_xl_path"]
        cls.column_to_idx: dict = config['column_to_idx']


    def get_gst_state(cls, gst_code: str)->str:
        try:
            # Get the absolute path to the Excel file
            excel_file_path = os.path.join( cls.gst_mapping_xl_path)
            
            # Read the Excel file into a DataFrame
            df = pd.read_excel(excel_file_path)
            
            # Check if the input_value exists in the 'Mapping from' column
            mapping_to = df.loc[df['TIN'] == gst_code, 'State'].values

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
            igst_amt = round(amount * gst_percentage / 100, 2) if calculate_igst else 0.0
            discount = 0.0
            total_amt = round(amount + (2 * cgst_sgst_amt) + igst_amt - discount, 2)
            new_row = cls.get_default_row_format(
                {
                    **{
                        i["column"]: str(row.iloc[i["target_column_idx"]]).strip()
                        for i in cls.target_columns_index
                    },
                    "Party's GST" : "" if gst_no == "nan" else gst_no,
                    "Reg Type" : "unregistered consumer" if gst_no == 'nan' else "regular",
                    "Product's Name": f"Medicine {gst_percentage}%",
                    "State" : "Uttrakhand" if gst_no == "nan" else state,
                    "Party/Cash" : str(row.iloc[cls.column_to_idx['Party/Cash']])[5:].strip(),
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

    def transform(cls, df: pd.DataFrame, save=True):
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

        filename = cls.get_filename_with_datetime(prefix="Purchase_")
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
            xl_save_path=xl_save_path,
            save_dir=cls.file_save_dir,
            xl_file_name = filename

        )


class TransformExcelSale(BaseTransformExcel):

    def __init__(cls, config: dict, *args, **kwargs) -> None:
        cls.output_columns = config["output_columns"]
        cls.gst_data = config["gst_data"]
        cls.default_output_row = config["default_output_row"]
        cls.columns_to_sum = config["columns_to_sum"]
        cls.target_columns_index = config["target_columns_index"]
        cls.file_save_dir = "transformed"
        cls.perfix_for_totals = config["perfix_for_totals"]
        cls.party_gst_index = config["party_gst_index"]
        
        # changes every time
        cls.bill_no_prefix = kwargs.get('bill_no_prefix', None)
        cls.bill_no_suffix_counter = kwargs.get('bill_no_suffix_counter', None)
        cls.calculate_igst = kwargs.get('calculate_igst', False)


    def get_transformed_rows(cls, row, bill_date: str,  bill_no: str, calculate_igst=False):
        new_rows = []
        gst_no = str(row.iloc[cls.party_gst_index])

        for gst in cls.gst_data:
            gst_percentage = gst["gst_percentage"]
            cgst_sgst_per = gst_percentage / 2
            amount = float(row.iloc[gst["target_column_idx"]])
            if amount == 0.0:
                continue

            cgst_sgst_amt =  round(amount * cgst_sgst_per / 100, 2)
            igst_amt = round(amount * gst_percentage / 100, 2) if calculate_igst else 0.0
            discount = 0.0
            total_amt = round(amount + (2 * cgst_sgst_amt) + igst_amt - discount, 2)
            new_row = cls.get_default_row_format(
                {
                    **{
                        i["column"]: str(row.iloc[i["target_column_idx"]]).strip()
                        for i in cls.target_columns_index
                    },
                    "Reg Type" : "unregistered consumer" if gst_no == 'nan' else "regular",
                    "Product's Name": f"Medicine {gst_percentage}%",
                    "Bill No." : bill_no,
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

    def transform(cls, df: pd.DataFrame, save=True):
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
                    bill_no = f"{cls.bill_no_prefix}{bill_counter:05d}",
                    calculate_igst =  cls.calculate_igst
                    )
                rows_to_df += new_tranformed_rows
                
                # if new rows added then only increase the bill counter
                # there may be case where no rows were added because amount is 0
                if len(new_tranformed_rows) > 0:
                    bill_counter += 1

        filename = cls.get_filename_with_datetime(prefix="Sale_")
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
            xl_save_path=xl_save_path,
            save_dir=cls.file_save_dir,
            xl_file_name = filename
        )


class TransformStockExcel(BaseTransformExcel):
    def __init__(cls, config: dict, *args, **kwargs) -> None:
        cls.file_save_dir = "transformed"

        pass

    def transform(cls, df: pd.DataFrame, save=True) -> pd.DataFrame:
        '''
        - Iterate over df rows
        - If row == header then append opening balalnce data.
        '''
        result = []
        stock_name = None
        for index, row in df.iterrows():
            # if the header found
            if row[0] == "Sr." and row[1] == "Transaction Date" and row[2] == "Transaction Type" and row[3] == "Amount" and row[4] == "Units":
                # if header found then get stock name from prv cell
                stock_name = df.iloc[index - 1, 0]
                # also append the next cell data to the result which is opening balance
                result.append([len(result)+1, stock_name, "", df.iloc[index + 1, 1], df.iloc[index + 1, 3], ""])
            # if the cell is not null and has int data at first
            elif pd.notnull(row[0]) and isinstance(row[0], int) and str(row[2]) != "nan":

                result.append([len(result)+1, stock_name, row[1], row[2], row[3], row[4]])
            # if nothing matches then continue
            else:
                continue

        columns = ["Sr.", "Fund Name", "Transaction Date", "Transaction Type", "Amount", "Units"]

        df = pd.DataFrame(result, columns=columns)


        if save:
            filename = cls.get_filename_with_datetime(prefix="Stock_")
            xl_save_path = os.path.join(cls.file_save_dir, filename)
            df.to_excel(xl_save_path, index=False)

        return dict(
            xl_save_path=xl_save_path,
            save_dir=cls.file_save_dir,
            xl_file_name = filename
        )



def check_for_stock():
    default_config = {}
    obj = TransformStockExcel(default_config)
    def get_df():
        sale_path = "/home/akshat/Documents/projects/joshi-uncle/data/stock.xlsx"
        df = pd.read_excel(
            sale_path, header=None
        )
        return df
    obj.transform(get_df())
    

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


if __name__ == "__main__":
    check_for_stock()
