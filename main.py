import json
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from datetime import datetime


class ExcelToJsonConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Conversor de Ensaios Balísticos")

        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=10, pady=10)

        self.btn_upload = tk.Button(
            self.frame, text="Carregar Excel", command=self.upload_excel
        )
        self.btn_upload.pack(side=tk.LEFT)

        self.btn_export = tk.Button(
            self.frame,
            text="Exportar JSON",
            command=self.export_json,
            state=tk.DISABLED,
        )
        self.btn_export.pack(side=tk.LEFT, padx=5)

        self.status_label = tk.Label(self.root, text="")
        self.status_label.pack(pady=5)

        self.data = {}
        self.current_file = ""

    def upload_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.current_file = file_path
            self.process_excel()
            self.btn_export.config(state=tk.NORMAL)
            self.status_label.config(text="Arquivo carregado com sucesso!")

    def process_excel(self):
        wb = load_workbook(self.current_file, data_only=True)
        self.data = {"temperaturas": {}}

        for sheet_name in wb.sheetnames:
            if "minus" in sheet_name:
                temp_type = "LT"
            elif "RT" in sheet_name:
                temp_type = "RT"
            elif "plus" in sheet_name:
                temp_type = "HT"
            else:
                continue

            if "Datenblatt" in sheet_name:
                self.process_datenblatt(wb[sheet_name], temp_type)
            elif "Grafik" in sheet_name:
                self.process_grafik(wb[sheet_name], temp_type)

    def process_datenblatt(self, sheet, temp_type):
        meta = {
            "test_order": self.clean_value(sheet["J4"].value),
            "inflator_type": self.clean_value(sheet["U1"].value),
            "production_order": self.clean_value(sheet["J3"].value),
            "propellant_lot_number": self.clean_value(sheet["S3"].value),
            "test_date": self.parse_date(sheet["C4"].value),
        }

        tests = []
        for row in sheet.iter_rows(min_row=10, values_only=True):
            if row[0] and str(row[0]).isdigit():
                tests.append(
                    {
                        "test_no": self.clean_value(row[0]),
                        "inflator_no": self.clean_value(row[1]),
                    }
                )

        if temp_type not in self.data["temperaturas"]:
            self.data["temperaturas"][temp_type] = {}

        self.data["temperaturas"][temp_type].update(
            {"metadados": meta, "ensaios": tests}
        )

    def process_grafik(self, sheet, temp_type):
        method = self.detect_method(sheet)
        pressure_data = []

        for row_idx, row in enumerate(sheet.iter_rows(min_row=60), start=60):
            if row[0].value and str(row[0].value).isdigit():
                inflator_no = self.clean_value(row[0].value)
                pressures = {}

                for col_idx, cell in enumerate(
                    row[2:150], start=3
                ):  # Colunas C em diante
                    if cell.value is not None:
                        ms = self.get_millisecond(sheet, row_idx, col_idx, method)
                        if ms is not None:
                            pressures[str(ms)] = float(cell.value)

                pressure_data.append(
                    {"inflator_no": inflator_no, "pressoes": pressures}
                )

        if temp_type in self.data["temperaturas"]:
            self.data["temperaturas"][temp_type]["dados_pressao"] = pressure_data

    def detect_method(self, sheet):
        if any(sheet.cell(row=51, column=col).value for col in range(1, 50)) or any(
            sheet.cell(row=55, column=col).value for col in range(1, 50)
        ):
            return "AKLV"
        return "USCAR"

    def get_millisecond(self, sheet, row, col, method):
        if method == "AKLV":
            min_col = sheet.cell(row=51, column=col).value
            max_col = sheet.cell(row=55, column=col).value
            return min_col if min_col else max_col
        else:
            if col >= 136 and col <= 150:  # Colunas EX (136) até FD (150)
                return sheet.cell(row=56, column=col).value

    def clean_value(self, value):
        if isinstance(value, str):
            return value.strip().lstrip("'")
        return value

    def parse_date(self, date_value):
        if isinstance(date_value, datetime):
            return date_value.strftime("%Y-%m-%d")
        return date_value

    def export_json(self):
        if self.current_file:
            output_file = self.current_file.replace(".xlsx", ".json")
            with open(output_file, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
            self.status_label.config(text=f"Arquivo exportado: {output_file}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    converter = ExcelToJsonConverter()
    converter.run()
