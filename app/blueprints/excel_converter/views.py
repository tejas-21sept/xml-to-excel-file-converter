import os
import xml.etree.ElementTree as ET
from datetime import datetime
from io import BytesIO

import pandas as pd
from flask import Response, jsonify, request
from flask.views import MethodView
from lxml import etree


class ConerterAPI(MethodView):

    def post(self):
        file = request.files.get("file")
        if not file:
            return jsonify({"message": "No file provided"}), 400

        if not file.filename.lower().endswith(".xml"):
            return (
                jsonify({"message": "Invalid file type. Please upload an XML file."}),
                400,
            )

        try:
            tree = etree.parse(file)
            root = tree.getroot()
            excel_data = self.parse_xml_to_csv(root)

            # Create a response with the Excel file as an attachment

            response = Response(
                excel_data,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response.headers["Content-Disposition"] = (
                "attachment; filename=tally_data.xlsx"
            )
            return response

        except etree.XMLSyntaxError:
            return jsonify({"message": "Error parsing XML file."}), 400
        except Exception as e:
            return jsonify({"message": f"Error: {str(e)}"}), 500

    def format_date(self, date_str):
        date_obj = datetime.strptime(date_str, "%Y%m%d")
        return date_obj.strftime("%d-%m-%Y")

    def values_in_cells(self, value):
        if value:
            return value
        elif value is None:
            return ""
        return "NA"

    def set_debtor_value(self, txn_type, tag, tag_value, row, column_mapping):
        try:
            print(f"\nset_debtor_value - {txn_type} || {tag} || {tag_value} || {row}")
            if tag.upper() == "PARTYLEDGERNAME" and txn_type == "Parent" and tag_value:
                row[column_mapping.get("debtor", "debtor")] = self.values_in_cells(
                    tag_value
                )
            elif (
                tag.upper() == "LEDGERNAME"
                and txn_type in ("Child", "Other")
                and tag_value
            ):
                row[column_mapping.get("debtor", "debtor")] = self.values_in_cells(
                    tag_value
                )
            row[column_mapping.get("particulars", "particulars")] = (
                self.values_in_cells(tag_value)
            )
            print(f"\ndebt row - {row} - {id(row)}\n")
            return True
        except Exception as e:
            print(f"\nset_debtor_value - {e} \n")
            return False

    def store_already_fetched_debtor(
        self, debtor_tag, debtor_tag_value, txn_type, row, column_mapping
    ):
        try:
            modified_debtor_tag = debtor_tag.upper()
            debtor_creation = self.set_debtor_value(
                txn_type, modified_debtor_tag, debtor_tag_value, row, column_mapping
            )
            return debtor_creation
        except Exception as e:
            print(f"\n store_already_fetched_debtor - {e}\n ")
            return False

    def has_multiple_billallocations_with_data(self, tallymessage):
        bill_allocations = tallymessage.findall(".//BILLALLOCATIONS.LIST")
        non_empty_allocations = []

        for allocation in bill_allocations:
            # Check if any child tag has meaningful data
            has_data = any(
                child.tag.upper() in ("NAME", "BILLTYPE", "AMOUNT")
                and child.text.strip()
                for child in allocation
            )
            if has_data:
                non_empty_allocations.append(allocation)
        print(f"\n non_empty_allocations - {len(non_empty_allocations)} \n")
        return non_empty_allocations

    def parse_xml_to_csv(self, root):

        # excel file column ordering
        column_mapping = {
            "DATE": "Date",
            "txn_type": "Transaction Type",
            "VOUCHERNUMBER": "Vch No.",
            "ref_no": "Ref No",
            "ref_type": "Ref Type",
            # "ref_date": "Ref Date",
            "debtor": "Debtor",
            "ref_amount": "Ref Amount",
            "amt": "Amount",
            "particulars": "Particulars",
            "VOUCHERTYPENAME": "Vch Type",
            "amt_verification": "Amount Verified",
        }

        # Initialize a list to hold rows for CSV
        rows = []

        # Process each tallymessage
        for tallymessage in root.findall(
            ".//TALLYMESSAGE/VOUCHER[@VCHTYPE='Receipt']/.."
        ):
            shared_row = {}
            rows_for_message = []
            tag_d = {}
            debtor_set = False
            debtor_set_values = dict()
            txn_type = None
            billallocations_count = self.has_multiple_billallocations_with_data(
                tallymessage
            )
            voucher_number = None
            voucher_type = None

            # Extract shared data first
            for element in tallymessage.iter():
                tag = element.tag
                tag_value = element.text.strip() if element.text else "Empty"
                tag_d[tag] = tag_value

                if tag.upper() == "DATE" and tag_value:
                    value = self.format_date(tag_value)
                    shared_row[column_mapping.get(tag, tag)] = self.values_in_cells(
                        value
                    )
                elif tag.upper() == "VOUCHERNUMBER" and tag_value:
                    print(f"\nvch no - {tag_value} \n")
                    shared_row[column_mapping.get(tag, tag)] = self.values_in_cells(
                        tag_value
                    )
                elif (
                    not debtor_set
                    and tag.upper() in ("PARTYLEDGERNAME", "LEDGERNAME")
                    and tag_value
                ):
                    if txn_type:
                        print(f"\ntxn_type")
                        debtor_creation = self.set_debtor_value(
                            txn_type, tag, tag_value, shared_row, column_mapping
                        )
                        debtor_set = debtor_creation
                    else:
                        debtor_set_values[tag.lower()] = tag_value
                elif tag.upper() == "VOUCHERTYPENAME" and tag_value:
                    voucher_type = tag_value
                    shared_row[column_mapping.get(tag, tag)] = self.values_in_cells(
                        tag_value
                    )

                elif tag.upper() == "VOUCHERNUMBER" and tag_value:
                    voucher_number = tag_value

            # Process BILLALLOCATIONS.LIST if present
            for allocation in self.has_multiple_billallocations_with_data(tallymessage):
                # if not any(child.tag.upper() in ("NAME", "BILLTYPE", "AMOUNT") and child.text.strip() for child in element ):
                #     continue

                row = shared_row.copy()
                ref_no = None
                ref_type = None

                for bill_data in allocation:
                    bill_type_tag = bill_data.tag
                    # print(f"\n bill_type_tag - {bill_type_tag}")

                    if bill_type_tag.upper() == "BILLTYPE" and bill_data.text:
                        # print(f"\nbill_type_tag text - {bill_data.text} \n")
                        ref_type = bill_data.text.strip()
                        modified_bill_type = ref_type.lower()
                        if modified_bill_type in ("agst ref", "new ref"):
                            txn_type = "Child"
                        elif modified_bill_type in ("bank", "gst"):
                            txn_type = "Other"
                        elif voucher_number and voucher_type:
                            txn_type = "Parent"

                    elif bill_type_tag.upper() == "NAME" and bill_data.text:
                        ref_no = bill_data.text

                    if (
                        txn_type == "Child"
                        and bill_type_tag.upper() == "AMOUNT"
                        and bill_data.text
                    ):
                        row[column_mapping.get("ref_amount", "ref_amount")] = (
                            self.values_in_cells(bill_data.text)
                        )
                        row[column_mapping.get("amt", "amt")] = self.values_in_cells("")
                    elif (
                        txn_type == "Parent"
                        and bill_type_tag.upper() == "AMOUNT"
                        and bill_data.text
                    ):
                        row[column_mapping.get("ref_amount", "ref_amount")] = (
                            self.values_in_cells("")
                        )
                        # need to fetch all child transaction and there addition
                        row[column_mapping.get("amt", "amt")] = self.values_in_cells(
                            bill_data.text
                        )

                    elif (
                        txn_type == "Other"
                        and bill_type_tag.upper() == "AMOUNT"
                        and bill_data.text
                    ):
                        row[column_mapping.get("ref_amount", "ref_amount")] = (
                            self.values_in_cells("")
                        )
                        row[column_mapping.get("amt", "amt")] = self.values_in_cells(
                            bill_data.text
                        )

                    if not debtor_set and len(debtor_set_values) > 1 and txn_type:
                        print(f"\n for loop debt")
                        debtor_tag = (
                            "partyledgername" if txn_type == "Parent" else "ledgername"
                        )
                        debtor_tag_value = debtor_set_values[debtor_tag]
                        print(
                            f"\n debtor_tag - {debtor_tag} || {debtor_tag_value} || {txn_type} || \n"
                        )
                        debtor_creation = self.store_already_fetched_debtor(
                            debtor_tag,
                            debtor_tag_value,
                            txn_type,
                            shared_row,
                            column_mapping,
                        )
                        debtor_set = debtor_creation

                # Add BILLALLOCATIONS.LIST-specific data
                row[column_mapping.get("ref_no", "ref_no")] = self.values_in_cells(
                    ref_no
                )
                row[column_mapping.get("txn_type", "txn_type")] = self.values_in_cells(
                    txn_type
                )
                row[column_mapping.get("ref_type", "ref_type")] = self.values_in_cells(
                    ref_type
                )

                rows_for_message.append(row)
                print(f"\n append row - {row} - {id(row)}\n")

            # Handle case with no BILLALLOCATIONS.LIST
            if len(billallocations_count) == 0:
                rows_for_message.append(shared_row)

            # Ensure each row respects the column order
            for row_data in rows_for_message:
                ordered_row = {
                    column_mapping[col]: row_data.get(column_mapping[col], "")
                    for col in column_mapping.keys()
                }
                rows.append(ordered_row)
                print(f"\nordered_row - {ordered_row} -- {debtor_set_values} \n")
                print("--" * 20)

        print(f"\n Rows len - {len(rows)}\n")
        print(f"\n Rows - {rows} \n")

        # Ensure consistency in headers (with custom column names)
        df = pd.DataFrame(rows)

        # Write to Excel file
        output_file = "output.xlsx"
        df.to_excel(output_file, index=False)
        print(f"Excel file {output_file} generated successfully.")

        df = pd.DataFrame(rows)

        excel_buffer = BytesIO()
        df.to_excel(excel_buffer, index=False, engine="openpyxl")
        excel_buffer.seek(0)

        return excel_buffer


# def parse_xml_to_csv(input_file, output_folder):
#     tree = ET.parse(input_file)
#     root = tree.getroot()

#     column_mapping = {
#         "DATE": "Date",
#         "txn_type": "Transaction Type",
#         "VOUCHERNUMBER": "Vch No.",
#         "ref_no": "Ref No",
#         "ref_type": "Ref Type",
#         "debtor": "Debtor",
#         "ref_amount": "Ref Amount",
#         "amt": "Amount",
#         "particulars": "Particulars",
#         "VOUCHERTYPENAME": "Vch Type",
#         "amt_verification": "Amount Verified",
#     }
#     rows = []

#     for tallymessage in root.findall(".//TALLYMESSAGE/VOUCHER[@VCHTYPE='Receipt']/.."):
#         shared_row = {}

#         for element in tallymessage.iter():
#             tag = element.tag
#             tag_value = element.text.strip() if element.text else "Empty"
#             if tag.upper() == "DATE" and tag_value:
#                 shared_row[column_mapping.get(tag, tag)] = format_date(tag_value)
#             elif tag.upper() == "VOUCHERNUMBER" and tag_value:
#                 shared_row[column_mapping.get(tag, tag)] = tag_value
#             elif tag.upper() == "VOUCHERTYPENAME" and tag_value:
#                 shared_row[column_mapping.get(tag, tag)] = tag_value

#         rows.append(shared_row)

#     df = pd.DataFrame(rows)
#     output_file = os.path.join(output_folder, "output.xlsx")
#     df.to_excel(output_file, index=False)

#     return output_file
