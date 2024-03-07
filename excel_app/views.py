from django.shortcuts import render
from django.http import HttpResponseRedirect, JsonResponse
import pandas as pd
from .utils import (
    TransformExcelSale,
    TransformExcelPurchase,
    TransformStockExcel,
    TransformExcelGST,
    GST2BTransfromation,
    TransformExcelJJOnly,
    IpdOpdTransfromation,
    EchsDueTransfromation,
    GSTR1Self,
    BaseTransformExcel,
)
import os
import platform


def index(request):
    return render(request, "excel_app/index.html")


def _open_file(file: str):
    # Get the current operating system
    current_os = platform.system()
    print(file)
    # Check the operating system and execute the appropriate command
    if current_os == "Windows":
        os.system(f"start {file}")  # Open file or folder on Windows
    elif current_os == "Linux":
        os.system(f"xdg-open {file}")  # Open file or folder on Linux
    else:
        print("Unsupported operating system")  # Handle unsupported


def open_excel(request):
    if request.method == "GET":
        file_to_open = request.GET.get("file_to_open")
        _open_file(os.path.join("transformed", file_to_open))

        return JsonResponse({"status": "OK", "code": 200})


def open_file(request):
    if request.method == "GET":
        file_to_open = request.GET.get("file_to_open")
        _open_file(file_to_open)

        return JsonResponse({"status": "OK", "code": 200})


def upload_excel(request):
    if request.method == "POST":
        excel_file = request.FILES["excel_file"]
        change_format = request.POST["change_format"]

        # df = pd.read_excel(excel_file)
        transform = None
        if change_format == "transform_sale":
            default_config = TransformExcelSale.read_config(
                os.path.join("excel_app", "config", "sale_config.json")
            )
            bill_no_prefix = request.POST["bill_no_prefix"]
            bill_no_suffix_counter = int(request.POST["bill_no_suffix_counter"])
            calculate_igst = (
                False if request.POST.get("caculate_igst", None) is None else True
            )

            transform = TransformExcelSale(
                default_config,
                bill_no_prefix=bill_no_prefix,
                bill_no_suffix_counter=bill_no_suffix_counter,
                calculate_igst=calculate_igst,
            )
            heading = "GSTR1 Transformation"
        elif change_format == "transform_purchase":
            global_constants = BaseTransformExcel.read_config(
                os.path.join("excel_app", "config", "constants.json")
            )
            default_config = TransformExcelPurchase.read_config(
                os.path.join("excel_app", "config", "purchase_config.json")
            )
            transform = TransformExcelPurchase(
                {
                    **default_config,
                    "gst_mapping_xl_path": global_constants[
                        "purchase_gst_codes_mapping_xl_path"
                    ],
                }
            )
            heading = "GSTR2 Transformation"

        elif change_format == "transform_stock":
            transform = TransformStockExcel({})
            heading = "Mutual Fund Transformation"

        elif change_format == "transform_gst":
            mapping_df = pd.read_excel(request.FILES["mapping_file"])

            default_config = TransformExcelGST.read_config(
                os.path.join("excel_app", "config", "gst_config.json")
            )
            bill_no_prefix = request.POST["bill_no_prefix"]
            bill_no_suffix_counter = int(request.POST["bill_no_suffix_counter"])

            transform = TransformExcelGST(
                default_config,
                bill_no_prefix=bill_no_prefix,
                bill_no_suffix_counter=bill_no_suffix_counter,
                mapping_df=mapping_df,
            )
            heading = "GST Transformation"

        elif change_format == "transform_jj-only":
            default_config = TransformExcelSale.read_config(
                os.path.join("excel_app", "config", "jj-only_config.json")
            )
            bill_no_prefix = request.POST["bill_no_prefix"]
            bill_no_suffix_counter = int(request.POST["bill_no_suffix_counter"])

            transform = TransformExcelJJOnly(
                default_config,
                bill_no_prefix=bill_no_prefix,
                bill_no_suffix_counter=bill_no_suffix_counter,
            )
            heading = "JJ/Only Transformation"

        elif change_format == "transform_ipd":
            transform = IpdOpdTransfromation(
                sort_columns="TPA", filename_prefix="IPD", _for="ipd"
            )
            heading = "IPD Transformation"

        elif change_format == "transform_opd":
            transform = IpdOpdTransfromation(
                sort_columns="Payment Mode", filename_prefix="OPD", _for="opd"
            )
            heading = "OPD Transformation"

        elif change_format == "transform_gst2b":
            default_config = TransformExcelSale.read_config(
                os.path.join("excel_app", "config", "gst2b_config.json")
            )
            transform = GST2BTransfromation(default_config)
            heading = "GST 2B Excel Transformation"
            excel_file = request.FILES.getlist("excel_file")

        elif change_format == "transform_EchsDue":
            default_config = TransformExcelSale.read_config(
                os.path.join("excel_app", "config", "echs_due-config.json")
            )
            transform = EchsDueTransfromation(default_config)
            heading = "ECHS DUE Transformation"
            excel_file = request.FILES.getlist("excel_file")

        elif change_format == "transform_gstr1self":
            default_config = TransformExcelSale.read_config(
                os.path.join("excel_app", "config", "gstr1_self_config.json")
            )
            calculate_igst = (
                False if request.POST.get("caculate_igst", None) is None else True
            )

            transform = GSTR1Self(
                default_config,
                calculate_igst=calculate_igst,
            )
            heading = "GSTR1 Self Transformation"

        data = transform.transform(excel_file, save=True)

        return render(
            request,
            "excel_app/display_dataframe.html",
            {"data": data, "heading": heading},
        )

    return render(request, "excel_app/index.html")
