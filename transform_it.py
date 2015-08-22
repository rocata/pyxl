#!/usr/bin/env python

import sys
import json
import argparse

import xlsxwriter


def parse_args():
    parser = argparse.ArgumentParser(
        description="Transform a JSON structured dataset into an XLSx document (spreadsheet)."
        )
    
    def json_type(path_to_file):
        try:
            with open(path_to_file) as json_file:
                return json.load(json_file)
        except IOError:
            parser.error("File '%s' does not exist!" % path_to_file)

    def file_type(path_to_file):
        # TODO
        # check if there already exists a file at this path
        return path_to_file

    parser.add_argument(
        "input",
        type=json_type,
        help="Path to input file containing the data (JSON)"
        )
    parser.add_argument(
        "output",
        type=file_type,
        help="Path to the output file of the script. Can be an existing file and its content WILL be overwritten"
        )

    args = parser.parse_args()

    return args


def write_row(tab, row_number, items_in_row):

    column_number = 0
    for row_elem in items_in_row:
        tab.write(row_number, column_number, row_elem)
        column_number += 1

if __name__ == "__main__":
    args = parse_args()
    cfg = vars(args)

    # Create the workbook
    workbook = xlsxwriter.Workbook(cfg.get("output"))

    # Initialize worksheets
    ws_tab_1 = workbook.add_worksheet("Products")
    ws_tab_2 = workbook.add_worksheet("AdditionalImages")
    ws_tab_3 = workbook.add_worksheet("Specials")
    ws_tab_4 = workbook.add_worksheet("Discounts")
    ws_tab_5 = workbook.add_worksheet("Rewards")
    ws_tab_6 = workbook.add_worksheet("ProductOptions")
    ws_tab_7 = workbook.add_worksheet("ProductOptionValues")
    ws_tab_8 = workbook.add_worksheet("ProductAttributes")

    # Header rows for each worksheet
    header_tab_1 = ["product_id", "name(en)", "categories", "sku", "upc", "ean", "jan", "isbn", "mpn", "location", "quantity", "model", "manufacturer", "image_name", "shipping", "price", "points", "date_added", "date_modified", "date_available", "weight", "weight_unit", "length", "width", "height", "length_unit", "status", "tax_class_id", "seo_keyword", "description(en)", "meta_title(en)", "meta_description(en)", "meta_keywords(en)", "stock_status_id", "store_ids", "layout", "related_ids", "tags(en)", "sort_order", "subtract", "minimum"]
    header_tab_2 = ["product_id", "image", "sort_order"]
    header_tab_3 = ["product_id", "customer_group", "priority", "price", "date_start", "date_end"]
    header_tab_4 = ["product_id", "customer_group", "quantity", "priority", "price", "date_start", "date_end"]
    header_tab_5 = ["product_id", "customer_group", "points"]
    header_tab_6 = ["product_id", "option", "default_option_value", "required"]
    header_tab_7 = ["product_id", "option", "option_value", "quantity", "subtract", "price", "price_prefix", "points", "points_prefix", "weight", "weight_prefix"]
    header_tab_8 = ["product_id", "attribute_group", "attribute", "text(en)"]

    # I'm writing the headers for each worksheet
    write_row(ws_tab_1, 0, header_tab_1)
    write_row(ws_tab_2, 0, header_tab_2)
    write_row(ws_tab_3, 0, header_tab_3)
    write_row(ws_tab_4, 0, header_tab_4)
    write_row(ws_tab_5, 0, header_tab_5)
    write_row(ws_tab_6, 0, header_tab_6)
    write_row(ws_tab_7, 0, header_tab_7)
    write_row(ws_tab_8, 0, header_tab_8)

    #
    # counters
    # :/
    row_tab_1 = 1
    row_tab_2 = 1
    row_tab_3 = 1
    row_tab_6 = 1
    row_tab_7 = 1

    #
    # TAB 1 | Products
    #
    # Insert a row for each product
    for prd in cfg.get("input", {}).get("data", []):
        items = [
            prd.get("product_id"),
            prd.get("name"),
            prd.get("categories"),
            prd.get("sku"),
            prd.get("upc"),
            prd.get("ean"),
            prd.get("jan"),
            prd.get("isbn"),
            prd.get("mpn"),
            prd.get("location"),
            prd.get("quantity"),
            prd.get("model"),
            prd.get("manufacturer"),
            prd.get("image_name"),
            prd.get("shipping"),
            prd.get("price"),
            prd.get("points", 0),
            prd.get("date_added"),
            prd.get("date_modified"),
            prd.get("date_available"),
            prd.get("weight"),
            prd.get("weight_unit"),
            prd.get("length"),
            prd.get("width"),
            prd.get("height"),
            prd.get("length_unit"),
            prd.get("status"),
            prd.get("tax_class_id"),
            prd.get("seo_keyword"),
            prd.get("description(en)"),
            prd.get("meta_title(en)"),
            prd.get("meta_description(en)"),
            prd.get("meta_keywords(en)"),
            prd.get("stock_status_id"),
            prd.get("store_ids"),
            prd.get("layout"),
            prd.get("related_ids"),
            prd.get("tags(en)"),
            prd.get("sort_order"),
            prd.get("subtract", "true"),
            prd.get("minimum", 1)
            ]

        write_row(
            ws_tab_1,
            row_tab_1,
            items
            )

        #
        # TAB 2 | AdditionalImages
        #
        # Add a row for each image contained in the list product["image"]
        for image in prd.get("image", []):
            fields = [
                prd.get("product_id"),
                image.get("image"),
                image.get("sort_order")
                ]
            write_row(
                ws_tab_2,
                row_tab_2,
                fields
                )
            row_tab_2 += 1

        #
        # TAB 3 | Specials
        #
        # Add a row to this worksheet if the flag product["special"] is set to "true" (string)
        if prd.get("special", "").lower() == "true":
            fields = [
                prd.get("product_id"),
                prd.get("special_customer_group"),
                prd.get("special_priority"),
                prd.get("special_price"),
                prd.get("special_date_start"),
                prd.get("special_date_end")
                ]
            write_row(
                ws_tab_3,
                row_tab_3,
                fields
                )

        #
        # TAB 4 | Discounts
        #
        # Leave this worksheet empty

        #
        # TAB 5 | Rewards
        #
        # Add a row for each product with defaulted values
        fields = [
            prd.get("product_id"),
            "Default",
            "0"
            ]
        write_row(
            ws_tab_5,
            row_tab_1,
            fields
            )

        #
        # TAB 6 | ProductOptions
        #
        # If there is a flag product["marime"] set to "true" (string) 
        # then add a row for each product with defaulted values
        if prd.get("marime", "").lower() == "true":
            fields = [
                prd.get("product_id"),
                "Marime",
                "",
                "true"
            ]
            write_row(
                ws_tab_6,
                row_tab_6,
                fields
                )
            row_tab_6 += 1

            #
            # TAB 7 | ProductOptionValues
            #
            # For each element from the list product["marimi"] add a row on this worksheet
            # option_value is the only field which may vary
            for marime in prd.get("marimi", []):
                fields = [
                    prd.get("product_id"),
                    "Marime",
                    marime.get("marime"),
                    "1",
                    "true",
                    "0.00",
                    "+",
                    "0",
                    "+",
                    "0.00",
                    "+"
                    ]
                write_row(
                    ws_tab_7,
                    row_tab_7,
                    fields
                    )
                row_tab_7 += 1

        row_tab_1 += 1

    #
    # TAB 8 | ProductAttributes
    #
    # Leave this worksheet empty

    sys.stderr.write(
        "Successfully went through %d products!" % (row_tab_1 - 1)
        )

    # Close and save the workbook =]
    # C'est la vie
    workbook.close()
