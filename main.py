import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import os
import uuid


class ExcelToJsonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to JSON Converter")
        self.root.geometry("400x200")

        # Create and pack GUI elements
        self.label = tk.Label(root, text="Select an Excel file to convert to JSON")
        self.label.pack(pady=10)

        self.select_button = tk.Button(
            root, text="Select Excel File", command=self.load_excel
        )
        self.select_button.pack(pady=5)

        self.convert_button = tk.Button(
            root,
            text="Convert to JSON",
            command=self.convert_to_json,
            state=tk.DISABLED,
        )
        self.convert_button.pack(pady=5)

        self.status_label = tk.Label(root, text="", wraplength=350)
        self.status_label.pack(pady=10)

        self.sheet_data = None

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                # Load all sheets
                self.sheet_data = pd.read_excel(file_path, sheet_name=None)
                self.convert_button.config(state=tk.NORMAL)
                self.status_label.config(text=f"Loaded: {os.path.basename(file_path)}")

                # Print sheet names and columns for debugging
                print("Loaded Excel file with sheets:")
                for sheet_name, df in self.sheet_data.items():
                    print(f"\nSheet: {sheet_name}")
                    print("Columns (with indices):")
                    for idx, col in enumerate(df.columns):
                        print(f"  {idx}: {col}")
                    print("First 5 rows of data:")
                    print(df.head(5))
                    if "Grafik" in sheet_name:
                        print("\nRow 51 (Minimum values):")
                        print(
                            df.iloc[50:51] if len(df) > 50 else "Row 51 not available"
                        )
                        print("\nRow 55 (Maximum values):")
                        print(
                            df.iloc[54:55] if len(df) > 54 else "Row 55 not available"
                        )
                        print("\nRows 60 onward (Time-series data, first 5 rows):")
                        if len(df) > 59:
                            selected_cols = (
                                [df.columns[2]]
                                + [df.columns[3]]
                                + [
                                    df.columns[i]
                                    for i in range(149, 170)
                                    if i < len(df.columns)
                                ]
                            )
                            print(df.iloc[59:64][selected_cols])
                        else:
                            print("Rows 60+ not available")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
                self.status_label.config(text="Error loading file")

    def convert_to_json(self):
        if self.sheet_data is None:
            messagebox.showwarning("Warning", "No Excel file loaded")
            return

        try:
            # Initialize JSON structure
            json_data = {
                "metadata": {
                    "document_type": "Manufacturer's Test Certificate",
                    "standards": ["DIN 50049 - 2.3", "EN 10 204"],
                    "manufacturer": {
                        "name": "TRW Automotive",
                        "division": "TRW Airbag Systems GmbH & Co. KG Werk Laage",
                        "address": "Daimler-Benz Allee 1, D-18299 Laage",
                        "focus": "Occupant Restraint Systems",
                    },
                    "test_details": {
                        "report_type": "datareport for inflator",
                        "inflator_type": "SPI-2 EVO V124",
                        "production_order": "1994825800",
                        "propellant_lot_number": "0745306603",
                        "test_date": "2025-04-22",
                        "test_order": "613532",
                        "propellant_type": "Ponte de Lima",
                        "propellant_mass_g": 20.5,
                        "tank_volume_l": 28.3,
                        "tank_type": "TANK 28.3l (PdL)",
                    },
                },
                "test_data": {
                    "low": {"tests": [], "statistics": {}, "time_series": []},
                    "ambient": {"tests": [], "statistics": {}, "time_series": []},
                    "high": {"tests": [], "statistics": {}, "time_series": []},
                },
            }

            # Define pressure and combustion keys
            pressure_keys = [
                f"pK{i}"
                for i in [
                    1,
                    2,
                    3,
                    4,
                    5,
                    6,
                    7,
                    8,
                    9,
                    10,
                    12,
                    14,
                    15,
                    16,
                    18,
                    20,
                    25,
                    30,
                    40,
                    50,
                    60,
                    80,
                    100,
                    120,
                    150,
                ]
            ]
            combustion_keys = [
                "pBk_bar",
                "tpBk_ms",
                "pFt_bar",
                "tpFt_ms",
                "ttfg_ms",
                "Gt_bar_ms",
                "tpK1_percent_ms",
                "tpK10_percent_ms",
                "tpK25_percent_ms",
                "tpK50_percent_ms",
                "tpK75_percent_ms",
                "tpK90_percent_ms",
            ]
            tank_keys = ["pKm2_bar", "d_pK_percent", "d_fso_percent"]

            # Map sheet names to JSON keys and default temperatures
            sheet_map = {
                "Datenblatt minus": {
                    "key": "low",
                    "default_temp": -35,
                    "grafik": "Grafik minus",
                },
                "Datenblatt RT": {
                    "key": "ambient",
                    "default_temp": 23,
                    "grafik": "Grafik RT",
                },
                "Datenblatt plus": {
                    "key": "high",
                    "default_temp": 90,
                    "grafik": "Grafik plus",
                },
            }

            # Map expected column names to actual column names in Datenblatt sheets
            column_map = {
                "test_no": "TestNumber",  # Update based on debugging output
                "inflator_no": "InflatorID",  # Update based on debugging output
                "pKmax_bar": "pKmax",  # Update based on debugging output
                "tpKmax_ms": "tpKmax",  # Update based on debugging output
                **{key: key for key in pressure_keys},  # Update if names differ
                **{key: key for key in combustion_keys},  # Update if names differ
                **{key: key for key in tank_keys},  # Update if names differ
            }

            # Temperature column name for Datenblatt
            temp_column = "Temp"  # Update based on debugging output

            # Column indices for Grafik sheets (based on Excel: C=2, EX=149, FL=169)
            ms_column = 2  # Column C
            inflator_start = 3  # Column D
            uscar_start = 149  # Column EX
            uscar_end = 169  # Column FL

            # Process sheets
            for sheet_name, temp_df in self.sheet_data.items():
                if sheet_name not in sheet_map and "Grafik" not in sheet_name:
                    print(
                        f"Skipping sheet '{sheet_name}' (not a Datenblatt or Grafik sheet)"
                    )
                    continue

                # Process Datenblatt sheets
                if sheet_name in sheet_map:
                    json_key = sheet_map[sheet_name]["key"]
                    default_temp = sheet_map[sheet_name]["default_temp"]
                    print(
                        f"\nProcessing Datenblatt sheet '{sheet_name}' for {json_key}: {len(temp_df)} rows"
                    )

                    # Drop rows with all NaN values
                    temp_df = temp_df.dropna(how="all")
                    print(f"After dropping empty rows: {len(temp_df)} rows")

                    # Check for temperature column
                    actual_temp = default_temp
                    if temp_column in temp_df.columns:
                        temp_df[temp_column] = (
                            temp_df[temp_column]
                            .astype(str)
                            .str.replace("°C", "")
                            .str.replace("C", "")
                            .str.strip()
                        )
                        temp_df[temp_column] = pd.to_numeric(
                            temp_df[temp_column], errors="coerce"
                        )
                        unique_temps = temp_df[temp_column].dropna().unique()
                        print(f"Temperature values in {sheet_name}: {unique_temps}")
                        if len(unique_temps) > 0:
                            actual_temp = float(unique_temps[0])
                    else:
                        print(
                            f"Warning: No '{temp_column}' column in {sheet_name}. Using default temperature {default_temp}."
                        )

                    # Process individual tests
                    tests = []
                    for _, row in temp_df.iterrows():
                        test = {
                            "test_no": str(row.get(column_map["test_no"], "")),
                            "inflator_no": str(row.get(column_map["inflator_no"], "")),
                            "temperature_C": actual_temp,
                            "pKmax_bar": float(row.get(column_map["pKmax_bar"], 0.0)),
                            "tpKmax_ms": float(row.get(column_map["tpKmax_ms"], 0.0)),
                            "pressure_measurements_bar": {
                                key: float(row.get(column_map[key], 0.0))
                                for key in pressure_keys
                            },
                            "combustion_data": {
                                key: float(row.get(column_map[key], 0.0))
                                for key in combustion_keys
                            },
                            "tank_data": {
                                key: float(row.get(column_map[key], 0.0))
                                for key in tank_keys
                            },
                        }
                        tests.append(test)
                    json_data["test_data"][json_key]["tests"] = tests

                    # Compute statistics from Datenblatt
                    stats = {
                        "count": len(temp_df),
                        "min": {
                            key: (
                                float(temp_df[column_map[key]].min())
                                if column_map[key] in temp_df
                                and not temp_df[column_map[key]].isna().all()
                                else 0.0
                            )
                            for key in pressure_keys
                        },
                        "mittel": {
                            key: (
                                float(temp_df[column_map[key]].mean())
                                if column_map[key] in temp_df
                                and not temp_df[column_map[key]].isna().all()
                                else 0.0
                            )
                            for key in pressure_keys
                        },
                        "max": {
                            key: (
                                float(temp_df[column_map[key]].max())
                                if column_map[key] in temp_df
                                and not temp_df[column_map[key]].isna().all()
                                else 0.0
                            )
                            for key in pressure_keys
                        },
                        "ug": {
                            key: float(temp_df.get(f"ug_{key}", 0.0))
                            for key in ["pK2", "pK5", "pK8", "pK10", "pK16", "pK20"]
                        },
                        "og": {
                            key: float(temp_df.get(f"og_{key}", 0.0))
                            for key in ["pK2", "pK5", "pK8", "pK10", "pK16", "pK20"]
                        },
                    }
                    json_data["test_data"][json_key]["statistics"] = stats

                # Process Grafik sheets
                if "Grafik" in sheet_name:
                    for datenblatt_name, config in sheet_map.items():
                        if config["grafik"] == sheet_name:
                            json_key = config["key"]
                            print(
                                f"\nProcessing Grafik sheet '{sheet_name}' for {json_key}"
                            )

                            # Extract minimum values (row 51)
                            min_values = {}
                            min_reference = None
                            if len(temp_df) > 50:
                                min_row = temp_df.iloc[50]
                                for col in temp_df.columns:
                                    value = pd.to_numeric(
                                        min_row.get(col), errors="coerce"
                                    )
                                    if (
                                        not pd.isna(value)
                                        and col != temp_df.columns[ms_column]
                                    ):
                                        min_values[col] = float(value)
                                        if (
                                            min_reference is None
                                            or value
                                            < min_values.get(
                                                min_reference, float("inf")
                                            )
                                        ):
                                            min_reference = col
                                print(f"Minimum values (row 51): {min_values}")
                                print(f"Minimum reference column: {min_reference}")
                            else:
                                print("Warning: Row 51 not available in Grafik sheet")

                            # Extract maximum values (row 55)
                            max_values = {}
                            max_reference = None
                            if len(temp_df) > 54:
                                max_row = temp_df.iloc[54]
                                for col in temp_df.columns:
                                    value = pd.to_numeric(
                                        max_row.get(col), errors="coerce"
                                    )
                                    if (
                                        not pd.isna(value)
                                        and col != temp_df.columns[ms_column]
                                    ):
                                        max_values[col] = float(value)
                                        if (
                                            max_reference is None
                                            or value
                                            > max_values.get(
                                                max_reference, float("-inf")
                                            )
                                        ):
                                            max_reference = col
                                print(f"Maximum values (row 55): {max_values}")
                                print(f"Maximum reference column: {max_reference}")
                            else:
                                print("Warning: Row 55 not available in Grafik sheet")

                            # Extract time-series data (row 60 onward)
                            time_series = []
                            uscar_stats = []
                            if len(temp_df) > 59:
                                time_df = temp_df.iloc[59:].dropna(how="all")
                                print(
                                    f"Time-series data (rows 60+): {len(time_df)} rows"
                                )
                                for _, row in time_df.iterrows():
                                    # Time-series for AKLV (columns D to EW)
                                    time_point = {
                                        "ms": float(
                                            row.get(temp_df.columns[ms_column], 0.0)
                                        ),
                                        "values": {
                                            col: float(row.get(col, 0.0))
                                            for col in temp_df.columns[
                                                inflator_start:uscar_start
                                            ]
                                            if not pd.isna(row.get(col))
                                        },
                                    }
                                    time_series.append(time_point)

                                    # USCAR statistics (columns EX to FL)
                                    uscar_point = {
                                        "ms": float(
                                            row.get(temp_df.columns[ms_column], 0.0)
                                        ),
                                        "values": {
                                            col: float(row.get(col, 0.0))
                                            for col in temp_df.columns[
                                                uscar_start : uscar_end + 1
                                            ]
                                            if not pd.isna(row.get(col))
                                        },
                                    }
                                    uscar_stats.append(uscar_point)
                            else:
                                print("Warning: Rows 60+ not available in Grafik sheet")

                            # Update statistics with Grafik data
                            json_data["test_data"][json_key]["statistics"].update(
                                {
                                    "min_values": min_values,
                                    "min_reference": min_reference,
                                    "max_values": max_values,
                                    "max_reference": max_reference,
                                    "uscar_stats": uscar_stats,
                                }
                            )
                            json_data["test_data"][json_key][
                                "time_series"
                            ] = time_series

            # Save JSON to file
            output_file = os.path.join(os.getcwd(), f"output_{uuid.uuid4()}.json")
            with open(output_file, "w", encoding="utf-8") as f:
                json.dump(json_data, f, indent=2, ensure_ascii=False)

            self.status_label.config(
                text=f"JSON saved to: {os.path.basename(output_file)}"
            )
            messagebox.showinfo(
                "Success", f"JSON file generated successfully at: {output_file}"
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert to JSON: {str(e)}")
            self.status_label.config(text=f"Error converting to JSON: {str(e)}")
            print(f"Error details: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToJsonApp(root)
    root.mainloop()
