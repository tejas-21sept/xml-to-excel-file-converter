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

            if not tree:
                return jsonify({"message": "Invalid XML structure."}), 400

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
        return non_empty_allocations

    def add_ref_no(self, shared_row, column_mapping, txn_type, value=""):
        if txn_type in ("Child"):
            return self.values_in_cells(value)
        return self.values_in_cells("")

    def add_ref_type(self, shared_row, column_mapping, txn_type, value=""):
        if txn_type in ("Child"):
            return self.values_in_cells(value)
        return self.values_in_cells("")

    def add_ref_date(self, shared_row, column_mapping, txn_type, value=""):
        if txn_type in ("Child"):
            return self.values_in_cells(value)
        return self.values_in_cells("")

    def add_ref_amt(self, shared_row, column_mapping, txn_type, value=""):
        if txn_type in ("Child"):
            return self.values_in_cells(value)
        return self.values_in_cells("")

    def add_amt(self, shared_row, column_mapping, txn_type, value=""):
        if txn_type in ("Child"):
            return self.values_in_cells("")
        return self.values_in_cells(value)

    def add_amt_verified(self, shared_row, column_mapping, txn_type, value=""):
        if txn_type in ("Parent"):
            return self.values_in_cells(value)
        return self.values_in_cells("")

    def parse_xml_to_csv(self, root):

        # Order Xlsx Columns
        column_mapping = {
            "DATE": "Date",
            "txn_type": "Transaction Type",
            "VOUCHERNUMBER": "Vch No.",  # work to find parent and his vch no. here
            "ref_no": "Ref No",
            "ref_type": "Ref Type",
            "ref_date": "Ref Date",
            "debtor": "Debtor",
            "ref_amount": "Ref Amount",
            "amt": "Amount",
            "particulars": "Particulars",
            "VOUCHERTYPENAME": "Vch Type",
            "amt_verification": "Amount Verified",
        }
        # Initialize a list to hold rows for CSV
        rows = []

        # Loop through Tally XML structure - Process each tallymessage
        for tallymessage in root.findall(
            ".//TALLYMESSAGE/VOUCHER[@VCHTYPE='Receipt']/.."
        ):
            shared_row = {}
            data_rows = []
            date = None
            voucher_number = None
            voucher_type = None
            debtor_set_values = dict()
            txn_type = None
            ref_date = None
            total_child_amt = 0
            txn_amount = 0

            # Extract common data first
            for element in tallymessage.iter():
                tag = element.tag
                tag_value = element.text.strip() if element.text else None

                if tag.upper() == "DATE" and tag_value:
                    date = self.format_date(tag_value)

                elif tag.upper() == "VOUCHERNUMBER" and tag_value:
                    voucher_number = tag_value
                elif tag.upper() in ("PARTYLEDGERNAME", "LEDGERNAME") and tag_value:
                    debtor_set_values[tag.lower()] = tag_value

                elif tag.upper() == "VOUCHERTYPENAME" and tag_value:
                    voucher_type = tag_value

                elif tag.upper() == "REFERENCEDATE" and tag_value:
                    ref_date = self.format_date(tag_value)

                elif tag.upper() == "AMOUNT" and tag_value:
                    txn_amount = float(tag_value)

            # Process each bill
            for allocation in self.has_multiple_billallocations_with_data(tallymessage):

                temp_bill_data = {}
                ref_no = None
                ref_type = None

                for bill_data in allocation:  # single bill

                    bill_type_tag = bill_data.tag
                    bill_tag_value = bill_data.text.strip() if bill_data.text else None

                    if bill_type_tag.upper() == "NAME" and bill_tag_value:
                        temp_bill_data["Ref No"] = self.values_in_cells(bill_tag_value)

                    elif bill_type_tag.upper() == "BILLTYPE" and bill_tag_value:
                        ref_type = bill_data.text.strip()
                        modified_bill_type = ref_type.lower()
                        if modified_bill_type in ("agst ref", "new ref"):
                            txn_type = "Child"
                        elif modified_bill_type in ("bank", "gst"):
                            txn_type = "Other"
                        else:
                            break
                        temp_bill_data["Transaction Type"] = self.values_in_cells(
                            txn_type
                        )

                        temp_bill_data["Ref Type"] = self.add_ref_type(
                            temp_bill_data, column_mapping, txn_type, ref_type
                        )

                        temp_bill_data["Ref Date"] = self.add_ref_date(
                            temp_bill_data, column_mapping, txn_type, ref_date
                        )

                    if bill_type_tag.upper() == "AMOUNT" and bill_tag_value:
                        total_child_amt += float(bill_tag_value)
                        temp_bill_data["Ref Amount"] = self.add_ref_amt(
                            temp_bill_data, column_mapping, txn_type, bill_tag_value
                        )
                    temp_bill_data["Amount"] = self.add_amt(
                        temp_bill_data, txn_type, bill_tag_value
                    )

                    temp_bill_data[column_mapping.get("debtor", "debtor")] = (
                        self.values_in_cells(debtor_set_values["ledgername"])
                        if "ledgername" in debtor_set_values
                        else self.values_in_cells(None)
                    )
                    temp_bill_data[column_mapping.get("particulars", "particulars")] = (
                        self.values_in_cells(debtor_set_values["ledgername"])
                        if "ledgername" in debtor_set_values
                        else self.values_in_cells(None)
                    )

                    temp_bill_data["Amount Verified"] = self.values_in_cells("")
                    temp_bill_data["Date"] = self.values_in_cells(date)
                    temp_bill_data["Vch No."] = self.values_in_cells(voucher_number)
                    temp_bill_data["Vch Type"] = self.values_in_cells(voucher_type)

                ordered_temp_bill_data = {
                    col: temp_bill_data[col] if col in temp_bill_data else ""
                    for col in column_mapping.values()
                }

                data_rows.append(ordered_temp_bill_data)

            if voucher_number and voucher_type:
                txn_type = "Parent"
                shared_row[column_mapping.get("date", "date")] = (
                    self.values_in_cells(date) if date else self.values_in_cells(None)
                )
                shared_row[column_mapping.get("VOUCHERNUMBER", "VOUCHERNUMBER")] = (
                    self.values_in_cells(voucher_number)
                )
                shared_row[column_mapping.get("VOUCHERTYPENAME", "VOUCHERTYPENAME")] = (
                    self.values_in_cells(voucher_type)
                )
                shared_row[column_mapping.get("txn_type", "txn_type")] = (
                    self.values_in_cells(txn_type)
                )

                shared_row[column_mapping.get("DATE", "Date")] = self.values_in_cells(
                    date
                )
                shared_row[
                    column_mapping.get("amt_verification", "amt_verification")
                ] = (
                    self.values_in_cells("Yes")
                    if total_child_amt + txn_amount == 0
                    else self.values_in_cells("No")
                )
                shared_row[column_mapping.get("VOUCHERTYPENAME", "VOUCHERTYPENAME")] = (
                    self.values_in_cells(voucher_type)
                )
                shared_row[column_mapping.get("VOUCHERNUMBER", "VOUCHERNUMBER")] = (
                    self.values_in_cells(voucher_number)
                )

                shared_row[column_mapping.get("ref_no", "Ref No")] = self.add_ref_no(
                    shared_row, column_mapping, txn_type
                )
                shared_row[column_mapping.get("ref_type", "ref_type")] = (
                    self.add_ref_type(shared_row, column_mapping, txn_type)
                )
                shared_row[column_mapping.get("ref_date", "ref_date")] = (
                    self.add_ref_date(shared_row, column_mapping, txn_type)
                )
                shared_row[column_mapping.get("ref_amount", "ref_amount")] = (
                    self.add_ref_amt(shared_row, column_mapping, txn_type)
                )
                shared_row[column_mapping.get("debtor", "debtor")] = (
                    self.values_in_cells(debtor_set_values["partyledgername"])
                    if "partyledgername" in debtor_set_values
                    else self.values_in_cells(None)
                )
                shared_row[column_mapping.get("particulars", "particulars")] = (
                    self.values_in_cells(debtor_set_values["partyledgername"])
                    if "partyledgername" in debtor_set_values
                    else self.values_in_cells(None)
                )
                shared_row[column_mapping.get("amt", "amt")] = self.add_amt(
                    shared_row, column_mapping, txn_type, total_child_amt
                )
                ordered_shared_row = {
                    col: shared_row[col] if col in shared_row else ""
                    for col in column_mapping.values()
                }
                data_rows.append(ordered_shared_row)

            # Ensure each row respects the column order
            for row_data in data_rows:
                ordered_row = {
                    column_mapping[col]: row_data.get(column_mapping[col], "")
                    for col in column_mapping.keys()
                }
                rows.append(ordered_row)

        # Ensure consistency in headers (with custom column names)
        df = pd.DataFrame(rows)

        excel_buffer = BytesIO()
        df.to_excel(excel_buffer, index=False, engine="openpyxl")
        excel_buffer.seek(0)

        return excel_buffer
