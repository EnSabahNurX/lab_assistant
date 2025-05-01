import json
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
from datetime import datetime
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np


class ExcelToJsonConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Conversor de Ensaios Balísticos")
        self.root.geometry("1200x700")
        self.root.minsize(1000, 600)
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=3)
        self.root.rowconfigure(0, weight=1)

        # Frame esquerdo (Database)
        self.frame_db = tk.Frame(self.root, relief="groove", padx=10, pady=10, bd=2)
        self.frame_db.grid(row=0, column=0, sticky="nsew")
        self.frame_db.columnconfigure(0, weight=1)
        self.frame_db.rowconfigure(5, weight=1)

        # Título Database
        self.label_database_title = tk.Label(
            self.frame_db, text="Database", font=("Helvetica", 14, "bold")
        )
        self.label_database_title.grid(row=0, column=0, sticky="w", pady=(0, 10))

        # Entrada ordens
        tk.Label(
            self.frame_db, text="Digite os números das ordens separados por vírgula:"
        ).grid(row=1, column=0, sticky="w")
        self.entry_orders = tk.Entry(self.frame_db)
        self.entry_orders.grid(row=2, column=0, sticky="ew", pady=5)

        # Botões processar e remover ordens digitadas lado a lado
        btns_frame = tk.Frame(self.frame_db)
        btns_frame.grid(row=3, column=0, sticky="ew", pady=5)
        btns_frame.columnconfigure((0, 1), weight=1)

        self.btn_process = tk.Button(
            btns_frame, text="Adicionar Ordens Digitadas", command=self.process_orders
        )
        self.btn_process.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.btn_remove_orders = tk.Button(
            btns_frame,
            text="Remover Ordens Digitadas",
            command=self.remove_orders_by_input,
        )
        self.btn_remove_orders.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        self.status_label = tk.Label(self.frame_db, text="", anchor="w", fg="green")
        self.status_label.grid(row=4, column=0, sticky="ew", pady=(0, 10))

        # Gerenciador de ordens
        self.frame_orders_manager = tk.Frame(
            self.frame_db, relief="groove", bd=2, padx=10, pady=10
        )
        self.frame_orders_manager.grid(row=5, column=0, sticky="nsew")
        self.frame_orders_manager.columnconfigure(0, weight=1)
        self.frame_orders_manager.rowconfigure(2, weight=1)

        tk.Label(self.frame_orders_manager, text="Ordens:").grid(
            row=0, column=0, sticky="w", padx=5, pady=5
        )

        # Paginação e filtros
        self.current_page = 1
        self.orders_per_page = 10
        self.total_pages = 1

        self.pagination_frame = tk.Frame(self.frame_orders_manager)
        self.pagination_frame.grid(row=1, column=0, sticky="ew", padx=5)
        self.pagination_frame.columnconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)

        tk.Label(self.pagination_frame, text="Itens por página:").grid(
            row=0, column=0, sticky="w"
        )
        self.page_selector = ttk.Combobox(
            self.pagination_frame, values=[5, 10, 20, 50], width=5, state="readonly"
        )
        self.page_selector.set(self.orders_per_page)
        self.page_selector.grid(row=0, column=1, sticky="w")
        self.page_selector.bind("<<ComboboxSelected>>", self.update_items_per_page)

        self.select_all_var = tk.BooleanVar()
        self.select_all_chk = tk.Checkbutton(
            self.pagination_frame,
            text="Selecionar Tudo",
            variable=self.select_all_var,
            command=self.toggle_select_all,
        )
        self.select_all_chk.grid(row=0, column=2, sticky="w", padx=10)

        self.version_var = tk.StringVar()
        self.version_combobox = ttk.Combobox(
            self.pagination_frame,
            textvariable=self.version_var,
            state="readonly",
            width=12,
        )
        self.version_combobox.grid(row=0, column=3, sticky="w", padx=10)
        self.version_combobox.bind("<<ComboboxSelected>>", self.on_version_filter)
        self.version_combobox["values"] = []
        self.version_combobox.set("Todas")

        self.nav_frame = tk.Frame(self.pagination_frame)
        self.nav_frame.grid(row=0, column=6, sticky="e")

        self.prev_btn = tk.Button(
            self.nav_frame, text="< Anterior", command=lambda: self.change_page(-1)
        )
        self.prev_btn.pack(side="left")

        self.page_info = tk.Label(self.nav_frame, text="Página 1/1")
        self.page_info.pack(side="left", padx=5)

        self.next_btn = tk.Button(
            self.nav_frame, text="Próxima >", command=lambda: self.change_page(1)
        )
        self.next_btn.pack(side="left")

        # Canvas e scrollbar para ordens
        self.orders_canvas = tk.Canvas(self.frame_orders_manager)
        self.orders_canvas.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        self.orders_scrollbar = tk.Scrollbar(
            self.frame_orders_manager,
            orient=tk.VERTICAL,
            command=self.orders_canvas.yview,
        )
        self.orders_scrollbar.grid(row=2, column=1, sticky="ns", pady=5)
        self.orders_canvas.configure(yscrollcommand=self.orders_scrollbar.set)

        self.orders_inner_frame = tk.Frame(self.orders_canvas)
        self.orders_canvas.create_window(
            (0, 0), window=self.orders_inner_frame, anchor="nw"
        )

        self.orders_inner_frame.bind(
            "<Configure>",
            lambda e: self.orders_canvas.configure(
                scrollregion=self.orders_canvas.bbox("all")
            ),
        )
        # Forçar scroll para o topo ao atualizar a lista
        self.orders_canvas.yview_moveto(0)
        self.orders_canvas.bind(
            "<Enter>",
            lambda e: self.orders_canvas.bind_all("<MouseWheel>", self._on_mousewheel),
        )
        self.orders_canvas.bind(
            "<Leave>", lambda e: self.orders_canvas.unbind_all("<MouseWheel>")
        )

        self.order_vars = {}
        self.order_checkbuttons = {}

        # Botões abaixo do canvas
        self.btn_frame = tk.Frame(self.frame_orders_manager)
        self.btn_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        self.btn_frame.columnconfigure((0, 1, 2), weight=1)

        self.btn_send_workspace = tk.Button(
            self.btn_frame, text="Enviar ao Workspace", command=self.send_to_workspace
        )
        self.btn_send_workspace.grid(row=0, column=0, sticky="ew", padx=5)

        self.btn_remove_workspace_orders = tk.Button(
            self.btn_frame,
            text="Remover Ensaios das Ordens Selecionadas",
            command=self.remove_workspace_orders_selected,
        )
        self.btn_remove_workspace_orders.grid(row=0, column=1, sticky="ew", padx=5)

        self.btn_clear_workspace = tk.Button(
            self.btn_frame, text="Limpar Workspace", command=self.clear_workspace
        )
        self.btn_clear_workspace.grid(row=0, column=2, sticky="ew", padx=5)

        # Workspace (direita)
        self.workspace_frame = tk.Frame(
            self.root, relief="groove", bd=2, padx=10, pady=10
        )
        self.workspace_frame.grid(row=0, column=1, sticky="nsew")
        self.workspace_frame.columnconfigure(0, weight=1)
        self.workspace_frame.rowconfigure(2, weight=1)

        # Frame para o botão Relatório alinhado à direita no topo do workspace
        btn_report_frame = tk.Frame(self.workspace_frame)
        btn_report_frame.grid(row=0, column=0, sticky="ew")
        btn_report_frame.columnconfigure(0, weight=1)

        self.btn_report = tk.Button(
            btn_report_frame, text="Relatório", command=self.show_report
        )
        self.btn_report.grid(row=0, column=1, sticky="e")

        self.workspace_title = tk.Label(
            self.workspace_frame, text="Workspace", font=("Helvetica", 14, "bold")
        )
        self.workspace_title.grid(row=0, column=0, sticky="w", pady=(0, 10))

        # Filtros temperatura e limitador
        filter_temp_frame = tk.Frame(self.workspace_frame)
        filter_temp_frame.grid(row=1, column=0, sticky="w", pady=5)

        tk.Label(filter_temp_frame, text="Filtrar Temperatura:").pack(
            side="left", padx=(0, 5)
        )
        self.filter_temperature_var = tk.StringVar()
        self.filter_temperature_combobox = ttk.Combobox(
            filter_temp_frame,
            textvariable=self.filter_temperature_var,
            state="readonly",
            values=["Todas", "RT", "LT", "HT"],
            width=10,
        )
        self.filter_temperature_combobox.pack(side="left")
        self.filter_temperature_combobox.set("Todas")

        tk.Label(filter_temp_frame, text="Limitador:").pack(side="left", padx=(10, 5))
        self.limit_var = tk.StringVar()
        limit_values = ["Todos"] + [str(i) for i in range(5, 205, 5)]
        self.limit_combobox = ttk.Combobox(
            filter_temp_frame,
            textvariable=self.limit_var,
            state="readonly",
            values=limit_values,
            width=8,
        )
        self.limit_combobox.pack(side="left")
        self.limit_combobox.set("Todos")

        self.btn_apply_filters = tk.Button(
            filter_temp_frame, text="Aplicar Filtros", command=self.apply_filters
        )
        self.btn_apply_filters.pack(side="left", padx=10)

        # Lista de resultados com scrollbars
        self.frame_results = tk.Frame(self.workspace_frame)
        self.frame_results.grid(row=2, column=0, sticky="nsew", pady=5)
        self.frame_results.columnconfigure(0, weight=1)
        self.frame_results.rowconfigure(0, weight=1)

        self.scrollbar_y = tk.Scrollbar(self.frame_results, orient=tk.VERTICAL)
        self.scrollbar_x = tk.Scrollbar(self.frame_results, orient=tk.HORIZONTAL)

        self.list_results = tk.Listbox(
            self.frame_results,
            width=120,
            height=25,
            yscrollcommand=self.scrollbar_y.set,
            xscrollcommand=self.scrollbar_x.set,
        )
        self.list_results.grid(row=0, column=0, sticky="nsew")

        self.scrollbar_y.config(command=self.list_results.yview)
        self.scrollbar_y.grid(row=0, column=1, sticky="ns")

        self.scrollbar_x.config(command=self.list_results.xview)
        self.scrollbar_x.grid(row=1, column=0, sticky="ew")

        # Inicializações
        self.json_file = "Data.json"
        self.excel_folder = r"H:\TEAMS\Inflator_Lab\0_Evaluations\vi"
        self.workspace_data = []

        self.update_orders_list()
        self.root.mainloop()

    def _on_mousewheel(self, event):
        self.orders_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def process_orders(self):
        try:
            orders_input = self.entry_orders.get().strip()
            if not orders_input:
                messagebox.showerror("Erro", "Por favor, insira os números das ordens.")
                return

            orders = [order.strip() for order in orders_input.split(",")]

            if os.path.exists(self.json_file):
                with open(self.json_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = {}

            files_to_process = []
            for order in orders:
                for file_name in os.listdir(self.excel_folder):
                    if file_name.startswith(order) and file_name.endswith(".xlsx"):
                        files_to_process.append(
                            os.path.join(self.excel_folder, file_name)
                        )

            if not files_to_process:
                messagebox.showerror(
                    "Erro", "Nenhum arquivo Excel encontrado para as ordens fornecidas."
                )
                return

            for file in files_to_process:
                self.process_excel(file, data)

            with open(self.json_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            self.entry_orders.delete(0, tk.END)
            self.status_label.config(text="Base de dados JSON atualizada com sucesso!")
            messagebox.showinfo(
                "Sucesso", "Arquivos Excel processados e JSON atualizado!"
            )
            self.update_orders_list()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar ordens: {str(e)}")

    def process_excel(self, file_path, data):
        wb = load_workbook(file_path, data_only=True)
        current_version = None
        current_order = None
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
                current_version, current_order = self.process_datenblatt(
                    wb[sheet_name], temp_type, data
                )
            elif "Grafik" in sheet_name and current_version and current_order:
                self.process_grafik(
                    wb[sheet_name], temp_type, data, current_version, current_order
                )

    def process_datenblatt(self, sheet, temp_type, data):
        inflator_type = self.clean_value(sheet["U1"].value)
        version = "V" + inflator_type.split("V")[-1]

        test_order = self.clean_value(sheet["J4"].value)
        production_order = self.clean_value(sheet["J3"].value)
        propellant_lot_number = self.clean_value(sheet["S3"].value)
        test_date = self.parse_date(sheet["C4"].value)
        temperature_c = self.clean_value(sheet["C10"].value)

        if version not in data:
            data[version] = {}

        if test_order not in data[version]:
            data[version][test_order] = {
                "metadados": {
                    "production_order": production_order,
                    "propellant_lot_number": propellant_lot_number,
                    "test_date": test_date,
                },
                "temperaturas": {},
            }

        if temp_type not in data[version][test_order]["temperaturas"]:
            data[version][test_order]["temperaturas"][temp_type] = {
                "temperatura_c": float(temperature_c) if temperature_c else None,
                "ensaios": [],
            }

        tests = []
        seen_tests = set()
        for row in sheet.iter_rows(min_row=10, values_only=True):
            if row[0] and str(row[0]).strip().isdigit():
                test_no = self.clean_value(row[0])
                inflator_no = self.clean_value(row[1])
                if test_no and inflator_no and test_no not in seen_tests:
                    tests.append(
                        {"test_no": int(test_no), "inflator_no": int(inflator_no)}
                    )
                    seen_tests.add(test_no)

        data[version][test_order]["temperaturas"][temp_type]["ensaios"] = tests
        return version, test_order

    def process_grafik(self, sheet, temp_type, data, current_version, current_order):
        valid_columns = []
        limits = {"maximos": {}, "minimos": {}}
        for col in range(3, 151):
            min_val = self.clean_value(sheet.cell(row=51, column=col).value)
            max_val = self.clean_value(sheet.cell(row=55, column=col).value)
            if min_val or max_val:
                valid_columns.append(col)
                ms = col - 2
                if min_val is not None:
                    try:
                        limits["minimos"][str(ms)] = float(min_val)
                    except ValueError:
                        pass
                if max_val is not None:
                    try:
                        limits["maximos"][str(ms)] = float(max_val)
                    except ValueError:
                        pass

        inflator_nos = [
            ensayo["inflator_no"]
            for ensayo in data[current_version][current_order]["temperaturas"][
                temp_type
            ]["ensaios"]
        ]

        pressure_data = []
        blank_line_count = 0
        row_idx = 60

        for inflator_no in inflator_nos:
            is_blank = True
            pressures = {}
            for col in valid_columns:
                pressure = self.clean_value(sheet.cell(row=row_idx, column=col).value)
                if pressure is not None:
                    try:
                        pressures[str(col - 2)] = float(pressure)
                        is_blank = False
                    except ValueError:
                        continue

            if is_blank:
                blank_line_count += 1
                if blank_line_count >= 2:
                    break
            else:
                blank_line_count = 0
                if pressures:
                    pressure_data.append(
                        {"inflator_no": inflator_no, "pressoes": pressures}
                    )

            row_idx += 1

        data[current_version][current_order]["temperaturas"][temp_type][
            "dados_pressao"
        ] = pressure_data
        data[current_version][current_order]["temperaturas"][temp_type][
            "limites"
        ] = limits

    def update_orders_list(self):
        # Limpar Checkbuttons existentes
        for widget in self.orders_inner_frame.winfo_children():
            widget.destroy()
        self.order_vars.clear()
        self.order_checkbuttons.clear()

        if not os.path.exists(self.json_file):
            return

        try:
            with open(self.json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Preencher combobox de versões
            versions = sorted(data.keys())
            if not hasattr(self, "version_combobox"):
                # Se não existe, crie (adicione isso na __init__ se já não existir)
                self.version_var = tk.StringVar()
                self.version_combobox = ttk.Combobox(
                    self.frame_orders_manager,
                    textvariable=self.version_var,
                    state="readonly",
                    values=["Todas"] + versions,
                    width=12,
                )
                self.version_combobox.pack()
                self.version_combobox.bind(
                    "<<ComboboxSelected>>", self.on_version_filter
                )
            else:
                self.version_combobox["values"] = ["Todas"] + versions

            # Manter seleção se possível
            current = self.version_combobox.get()
            if current not in self.version_combobox["values"]:
                self.version_combobox.set("Todas")

            selected_version = self.version_combobox.get()

            # Coletar ordens com versão e test_date, filtrando pela versão selecionada
            orders_list = []
            for version, orders in data.items():
                if (
                    selected_version
                    and selected_version != "Todas"
                    and version != selected_version
                ):
                    continue
                for order, details in orders.items():
                    test_date = details["metadados"].get("test_date", "0000-00-00")
                    orders_list.append((version, order, test_date))

            # Ordenar por test_date (mais recente primeiro)
            def parse_date_safe(date_str):
                try:
                    return datetime.strptime(date_str, "%Y-%m-%d")
                except Exception:
                    return datetime.min

            orders_list.sort(key=lambda x: parse_date_safe(x[2]), reverse=True)

            # Adicionar Checkbuttons para cada ordem filtrada
            for idx, (version, order, test_date) in enumerate(orders_list):
                var = tk.BooleanVar()
                self.order_vars[(version, order)] = var
                display_text = f"Versão: {version}, Ordem: {order}, Data: {test_date}"

                # Frame para alinhar Checkbutton e botão de visualização
                row_frame = tk.Frame(self.orders_inner_frame)
                row_frame.grid(row=idx, column=0, sticky="w", padx=5, pady=2)

                chk = tk.Checkbutton(
                    row_frame,
                    text=display_text,
                    variable=var,
                    anchor="w",
                    width=45,
                )
                chk.pack(side=tk.LEFT)

                btn_view = tk.Button(
                    row_frame,
                    text="     👁️",
                    width=3,
                    command=lambda v=version, o=order: self.show_metadata_popup(v, o),
                )
                btn_view.pack(side=tk.LEFT, padx=(10, 0))

                self.order_checkbuttons[(version, order)] = chk

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar ordens: {str(e)}")

    def show_metadata_popup(self, version, order):
        # Carrega os dados do JSON
        if not os.path.exists(self.json_file):
            messagebox.showerror("Erro", "Base de dados não encontrada.")
            return

        with open(self.json_file, "r", encoding="utf-8") as f:
            data = json.load(f)

        if version not in data or order not in data[version]:
            messagebox.showerror("Erro", "Ordem não encontrada na base de dados.")
            return

        metadados = data[version][order].get("metadados", {})
        temperaturas = data[version][order].get("temperaturas", {})

        # Monta a string de metadados
        info = ""
        for k, v in metadados.items():
            info += f"{k}: {v}\n"
        info += "\nTemperaturas (°C):\n"
        for tipo, tdata in temperaturas.items():
            info += f"  {tipo}: {tdata.get('temperatura_c', 'N/A')}\n"

        # Cria a mini janela de visualização
        popup = tk.Toplevel(self.root)
        popup.title(f"Metadados - {order}")
        popup.geometry("350x250")
        popup.resizable(False, False)
        tk.Label(
            popup,
            text=f"Metadados da Ordem {order} ({version})",
            font=("Arial", 11, "bold"),
        ).pack(pady=8)
        text = tk.Text(popup, width=40, height=10, wrap="word")
        text.insert("1.0", info)
        text.config(state="disabled")
        text.pack(padx=8, pady=8)
        tk.Button(popup, text="Fechar", command=popup.destroy).pack(pady=5)

    def update_items_per_page(self, event=None):
        try:
            self.orders_per_page = int(self.page_selector.get())
            self.current_page = 1
            self.update_orders_list()
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Erro ao atualizar itens por página: {str(e)}"
            )

    def change_page(self, delta):
        new_page = self.current_page + delta
        if 1 <= new_page <= self.total_pages:
            self.current_page = new_page
            self.update_orders_list()

    def toggle_select_all(self):
        state = self.select_all_var.get()
        for var in self.order_vars.values():
            var.set(state)

    def on_version_filter(self, event=None):
        self.current_page = 1
        self.update_orders_list()

    def remove_orders_by_input(self):
        try:
            orders_input = self.entry_orders.get().strip()
            if not orders_input:
                messagebox.showerror("Erro", "Por favor, insira os números das ordens.")
                return

            orders_to_remove = [order.strip() for order in orders_input.split(",")]

            if not os.path.exists(self.json_file):
                messagebox.showerror("Erro", "Nenhum banco de dados encontrado.")
                return

            with open(self.json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            removed = []
            for version in list(data.keys()):
                for order in list(data[version].keys()):
                    if order in orders_to_remove:
                        del data[version][order]
                        removed.append(order)
                        if not data[version]:
                            del data[version]

            with open(self.json_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            if removed:
                msg = f"Ordens removidas com sucesso:\n{', '.join(removed)}"
                self.status_label.config(text=msg)
                messagebox.showinfo("Sucesso", msg)
            else:
                messagebox.showwarning(
                    "Aviso", "Nenhuma ordem correspondente encontrada."
                )

            self.entry_orders.delete(0, tk.END)
            self.update_orders_list()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao remover ordens: {str(e)}")

    def on_version_filter(self, event=None):
        """Atualiza a lista quando o filtro de versão muda"""
        self.update_orders_list()

    def send_to_workspace(self):
        """Envia dados detalhados dos ensaios selecionados para o Workspace."""
        selected_orders = [
            (version, order)
            for (version, order), var in self.order_vars.items()
            if var.get()
        ]
        if not selected_orders:
            messagebox.showwarning(
                "Aviso", "Nenhuma ordem selecionada para enviar ao Workspace."
            )
            return

        try:
            with open(self.json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Verificar versão dos ensaios selecionados
            versions = set(version for version, order in selected_orders)
            if len(versions) > 1:
                messagebox.showerror("Erro", "Selecione ensaios de apenas uma versão.")
                return

            # Se o Workspace já contém dados, verificar a versão
            if (
                self.workspace_data
                and self.workspace_data[0]["version"] not in versions
            ):
                messagebox.showerror(
                    "Erro",
                    "O Workspace já contém ensaios de outra versão. Limpe o Workspace antes.",
                )
                return

            new_workspace_data = []  # Dados temporários para adicionar

            for version, order in selected_orders:
                if version in data and order in data[version]:
                    details = data[version][order]
                    metadados = details.get("metadados", {})
                    test_date = metadados.get(
                        "test_date", "0000-00-00"
                    )  # Captura a data
                    temperaturas = details.get("temperaturas", {})
                    for temp_type, temp_data in temperaturas.items():
                        temperatura_c = temp_data.get("temperatura_c", "N/A")
                        ensaios = temp_data.get("ensaios", [])
                        dados_pressao = temp_data.get("dados_pressao", [])
                        pressao_map = {
                            item["inflator_no"]: item["pressoes"]
                            for item in dados_pressao
                        }
                        for ensaio in ensaios:
                            test_no = ensaio.get("test_no", "N/A")
                            inflator_no = ensaio.get("inflator_no", "N/A")
                            has_pressure_data = inflator_no in pressao_map
                            new_workspace_data.append(
                                {
                                    "test_no": test_no,
                                    "inflator_no": inflator_no,
                                    "temperatura_c": temperatura_c,
                                    "tipo": temp_type,
                                    "version": version,
                                    "order": order,
                                    "test_date": test_date,  # Inclui a data
                                    "pressoes": pressao_map.get(inflator_no, None),
                                }
                            )

            # Ordenar dados novos e existentes
            def parse_date_safe(date_str):
                try:
                    return datetime.strptime(date_str, "%Y-%m-%d")
                except:
                    return datetime(1900, 1, 1)  # Data mínima

            # Juntar novos registros aos existentes
            self.workspace_data.extend(new_workspace_data)
            self.workspace_data.sort(
                key=lambda x: parse_date_safe(x.get("test_date", "1900-01-01")),
                reverse=True,
            )

            self.update_workspace_display()  # Atualiza exibição

            messagebox.showinfo("Sucesso", "Ensaios enviados ao Workspace com sucesso!")

        except Exception as e:
            messagebox.showerror(
                "Erro", f"Erro ao enviar ensaios ao Workspace: {str(e)}"
            )

    def update_workspace_display(self):
        """Atualiza a exibição do Workspace com os dados atuais, ordenados por data."""
        self.list_results.delete(0, tk.END)
        header = "Ensaio | Inflator | Temperatura | Tipo | Versão | Ordem | Data"
        self.list_results.insert(tk.END, header)
        self.list_results.insert(tk.END, "-" * len(header))

        for reg in self.workspace_data:
            line = f"{reg['test_no']} | {reg['inflator_no']} | {reg['temperatura_c']}°C | {reg['tipo']} | {reg['version']} | {reg['order']} | {reg['test_date']}"
            if reg["pressoes"]:
                line += " | Dados de pressão disponíveis"
            else:
                line += " | Sem dados de pressão"
            self.list_results.insert(tk.END, line)

    def apply_filters(self):
        temp_filter = self.filter_temperature_var.get()
        limit_filter = self.limit_var.get()

        self.list_results.delete(0, tk.END)
        header = "Ensaio | Inflator | Temperatura | Tipo | Versão | Ordem"
        self.list_results.insert(tk.END, header)
        self.list_results.insert(tk.END, "-" * len(header))

        # Verificar se há múltiplas versões no workspace
        versions = {reg["version"] for reg in self.workspace_data}
        if len(versions) > 1:
            messagebox.showerror(
                "Erro",
                "Workspace contém versões mistas! Limpe antes de aplicar filtros.",
            )
            return

        filtered_data = []
        for reg in self.workspace_data:
            if temp_filter != "Todas" and reg["tipo"] != temp_filter:
                continue
            filtered_data.append(reg)

        # Aplicar limitador
        if limit_filter != "Todos":
            limit = int(limit_filter)
            filtered_data = filtered_data[:limit]

        # Exibir resultados filtrados
        for reg in filtered_data:
            line = f"{reg['test_no']} | {reg['inflator_no']} | {reg['temperatura_c']}°C | {reg['tipo']} | {reg['version']} | {reg['order']}"
            if reg["pressoes"]:
                line += " | Dados de pressão disponíveis"
            else:
                line += " | Sem dados de pressão"
            self.list_results.insert(tk.END, line)

    def remove_workspace_orders_selected(self):
        selected_orders = {
            (version, order)
            for (version, order), var in self.order_vars.items()
            if var.get()
        }
        if not selected_orders:
            messagebox.showwarning(
                "Aviso", "Nenhuma ordem selecionada para remover do Workspace."
            )
            return

        before = len(self.workspace_data)
        self.workspace_data = [
            reg
            for reg in self.workspace_data
            if (reg["version"], reg["order"]) not in selected_orders
        ]
        after = len(self.workspace_data)
        self.update_workspace_display()
        messagebox.showinfo(
            "Sucesso", f"{before - after} registros removidos do Workspace."
        )

    def clear_workspace(self):
        self.workspace_data.clear()
        self.list_results.delete(0, tk.END)
        messagebox.showinfo("Sucesso", "Workspace limpo com sucesso.")

    def clean_value(self, value):
        if value is None:
            return None
        if isinstance(value, str):
            return value.strip()
        return value

    def parse_date(self, date_value):
        if isinstance(date_value, datetime):
            return date_value.strftime("%Y-%m-%d")
        if isinstance(date_value, str):
            try:
                dt = datetime.strptime(date_value, "%d.%m.%Y")
                return dt.strftime("%Y-%m-%d")
            except Exception:
                return date_value
        return None

    def show_report(self):
        if not self.workspace_data:
            messagebox.showwarning(
                "Aviso", "Workspace vazio. Adicione ensaios antes de gerar o relatório."
            )
            return

        # Agrupar dados por temperatura
        data_by_temp = {}
        for reg in self.workspace_data:
            temp = reg["tipo"]
            if temp not in data_by_temp:
                data_by_temp[temp] = []
            data_by_temp[temp].append(reg)

        # Criar janela do relatório
        report_win = tk.Toplevel(self.root)
        report_win.title("Relatório de Ensaios")
        report_win.geometry("1000x700")

        # Frame com scrollbar vertical
        container = ttk.Frame(report_win)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Para cada temperatura, plotar gráfico e tabela
        for temp, records in data_by_temp.items():
            # Extrair versão (assumimos mesma versão para todos)
            versoes = set(r["version"] for r in records)
            if len(versoes) > 1:
                versao = ", ".join(versoes)
            else:
                versao = list(versoes)[0]

            # Número total de inflators
            total_inflators = len(records)

            # Criar frame para temperatura
            temp_frame = ttk.LabelFrame(
                scrollable_frame,
                text=f"Temperatura: {temp} | Versão: {versao} | Total Inflators: {total_inflators}",
            )
            temp_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            # Determinar pontos de tempo (ms) e montar matrizes para pressões
            # Consideramos que todos os registros têm dados_pressao: dict ms->pressao
            # Construir lista ordenada de ms (pontos de tempo)
            ms_points = set()
            for r in records:
                if r["pressoes"]:
                    ms_points.update(r["pressoes"].keys())
            ms_points = sorted(int(ms) for ms in ms_points)
            ms_points_str = [str(ms) for ms in ms_points]

            if not ms_points:
                ttk.Label(
                    temp_frame,
                    text="Nenhum dado de pressão disponível para esta temperatura.",
                ).pack()
                continue

            # Matriz pressões: linhas = inflators, colunas = ms_points
            pressao_matrix = []
            for r in records:
                p = []
                if r["pressoes"]:
                    for ms in ms_points_str:
                        val = r["pressoes"].get(ms, np.nan)
                        p.append(val)
                else:
                    p = [np.nan] * len(ms_points)
                pressao_matrix.append(p)
            pressao_matrix = np.array(pressao_matrix, dtype=np.float64)

            # Calcular limites maximos e minimos (pegando do primeiro registro, pois limites são iguais para todos)
            # Buscamos no JSON original: data[versao][ordem]["temperaturas"][temp]["limites"]
            limites_max = []
            limites_min = []
            try:
                with open(self.json_file, "r", encoding="utf-8") as f:
                    data_json = json.load(f)
                # Pegamos a ordem do primeiro registro
                ordem_exemplo = records[0]["order"]
                limites = data_json[versao][ordem_exemplo]["temperaturas"][temp][
                    "limites"
                ]
                max_dict = limites.get("maximos", {})
                min_dict = limites.get("minimos", {})
                limites_max = [max_dict.get(str(ms), np.nan) for ms in ms_points]
                limites_min = [min_dict.get(str(ms), np.nan) for ms in ms_points]
            except Exception:
                # Caso não encontre limites, preencher com nan
                limites_max = [np.nan] * len(ms_points)
                limites_min = [np.nan] * len(ms_points)

            # Calcular média ignorando nan
            media = np.nanmean(pressao_matrix, axis=0)

            # Criar figura matplotlib
            fig, ax = plt.subplots(figsize=(8, 4))
            # Curvas de todos inflators em preto
            for p in pressao_matrix:
                ax.plot(ms_points, p, color="black", linewidth=0.7, alpha=0.7)
            # Limite máximo em vermelho
            ax.plot(
                ms_points, limites_max, color="red", linewidth=2, label="Limite Máximo"
            )
            # Limite mínimo em vermelho
            ax.plot(
                ms_points, limites_min, color="red", linewidth=2, label="Limite Mínimo"
            )
            # Média em verde
            ax.plot(ms_points, media, color="green", linewidth=2, label="Média")

            ax.set_title(f"Curvas de Pressão - Temperatura {temp}")
            ax.set_xlabel("Tempo (ms)")
            ax.set_ylabel("Pressão (bar)")
            ax.legend()

            # Inserir figura no tkinter
            canvas_fig = FigureCanvasTkAgg(fig, master=temp_frame)
            canvas_fig.draw()
            canvas_fig.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            # Criar tabela abaixo do gráfico
            # Tabela com 3 linhas: limites máximos, média, limites mínimos
            from tkinter import ttk

            table = ttk.Treeview(
                temp_frame, columns=ms_points_str, show="headings", height=3
            )
            for ms in ms_points_str:
                table.heading(ms, text=ms)
                table.column(ms, width=30, anchor="center")

            # Inserir linhas
            def format_row(row):
                return [f"{v:.2f}" if not np.isnan(v) else "-" for v in row]

            table.insert("", "end", values=format_row(limites_max), tags=("max",))
            table.insert("", "end", values=format_row(media), tags=("media",))
            table.insert("", "end", values=format_row(limites_min), tags=("min",))

            table.tag_configure("max", background="#fdd")
            table.tag_configure("media", background="#dfd")
            table.tag_configure("min", background="#fdd")

            table.pack(fill=tk.BOTH, expand=True, pady=5)

        report_win.mainloop()


if __name__ == "__main__":
    ExcelToJsonConverter()
