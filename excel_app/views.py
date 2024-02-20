from django.shortcuts import render
from django.http import HttpResponseRedirect, JsonResponse
import pandas as pd
from .utils import TransformExcelSale, TransformExcelPurchase, TransformStockExcel,BaseTransformExcel
import webbrowser
import os
import platform

def index(request):
    return render(request, "excel_app/index.html")

def open_file(request):
    if request.method == "GET":
        file_to_open = request.GET.get("file_to_open")

        # Get the current operating system
        current_os = platform.system()
        print(file_to_open)
        # Check the operating system and execute the appropriate command
        if current_os == "Windows":
            os.system(f"start {file_to_open}")  # Open file or folder on Windows
        elif current_os == "Linux":
            os.system(f"xdg-open {file_to_open}")  # Open file or folder on Linux
        else:
            print("Unsupported operating system")  # Handle unsupported


        return JsonResponse({"status": "OK", "code": 200})


def upload_excel(request):
    if request.method == "POST":
        excel_file = request.FILES["excel_file"]
        change_format = request.POST["change_format"]
        df = pd.read_excel(excel_file)
        transform = None
        if change_format == "transform_sale":
            default_config = TransformExcelSale.read_config(
                os.path.join("excel_app", "config", "sale_config.json")
            )
            bill_no_prefix = request.POST["bill_no_prefix"]
            bill_no_suffix_counter = int(request.POST["bill_no_suffix_counter"])
            calculate_igst = False if request.POST.get("caculate_igst", None) is None else True

            transform = TransformExcelSale(
                default_config,
                bill_no_prefix=bill_no_prefix,
                bill_no_suffix_counter=bill_no_suffix_counter,
                calculate_igst=calculate_igst,
            )
            heading = "Sale Transformation"
        elif change_format == "transform_purchase":
            global_constants = BaseTransformExcel.read_config(
                os.path.join("excel_app", "config", "constants.json")
            )
            default_config = TransformExcelPurchase.read_config(
                os.path.join("excel_app", "config", "purchase_config.json")
            )
            transform = TransformExcelPurchase({**default_config, "gst_mapping_xl_path" : global_constants['purchase_gst_codes_mapping_xl_path']})
            heading = "Purchase Transformation"
            
        elif change_format == "transform_stock":
            transform = TransformStockExcel({})
            heading = "Stock Transformation"


        data = transform.transform(df, save=True)

        return render(
            request,
            "excel_app/display_dataframe.html",
            {"data": data, "heading": heading},
        )

    return render(request, "excel_app/index.html")
