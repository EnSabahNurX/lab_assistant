import json
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
from datetime import datetime
import os


class ExcelToJsonConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Conversor de Ensaios Balísticos")
        self.root.geometry("1200x600")
        self.root.resizable(True, True)

        # Variáveis para paginação
        self.current_page = 1
        self.orders_per_page = 10
        self.total_pages = 1

        # Seção para alimentar o banco de dados
        self.frame_db = tk.Frame(self.root)
        self.frame_db.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.Y)

        self.label_orders = tk.Label(
            self.frame_db, text="Digite os números das ordens separados por vírgula:"
        )
        self.label_orders.pack()

        self.entry_orders = tk.Entry(self.frame_db, width=50)
        self.entry_orders.pack()

        self.btn_process = tk.Button(
            self.frame_db,
            text="Processar Ordens e Atualizar JSON",
            command=self.process_orders,
        )
        self.btn_process.pack(pady=5)

        self.status_label = tk.Label(self.frame_db, text="")
        self.status_label.pack(pady=5)

        # Gerenciador de ordens
        self.frame_orders_manager = tk.Frame(self.frame_db)
        self.frame_orders_manager.pack(pady=10, fill=tk.BOTH, expand=True)

        self.label_orders_manager = tk.Label(self.frame_orders_manager, text="Ordens:")
        self.label_orders_manager.pack(anchor="w")

        # Frame para controles de paginação e seleção
        self.pagination_frame = tk.Frame(self.frame_orders_manager)
        self.pagination_frame.pack(fill=tk.X, pady=5)

        # Seletor de itens por página
        tk.Label(self.pagination_frame, text="Itens por página:").pack(
            side=tk.LEFT, padx=(0, 5)
        )
        self.page_selector = ttk.Combobox(
            self.pagination_frame,
            values=[10, 20, 30, 40, 50],
            width=5,
            state="readonly",
        )
        self.page_selector.set(self.orders_per_page)
        self.page_selector.pack(side=tk.LEFT)
        self.page_selector.bind("<<ComboboxSelected>>", self.update_items_per_page)

        # Checkbox selecionar tudo da página
        self.select_all_var = tk.BooleanVar()
        self.select_all_chk = tk.Checkbutton(
            self.pagination_frame,
            text="Selecionar Tudo",
            variable=self.select_all_var,
            command=self.toggle_select_all,
        )
        self.select_all_chk.pack(side=tk.LEFT, padx=10)

        # Navegação entre páginas
        self.nav_frame = tk.Frame(self.pagination_frame)
        self.nav_frame.pack(side=tk.RIGHT)

        self.prev_btn = tk.Button(
            self.nav_frame, text="< Anterior", command=lambda: self.change_page(-1)
        )
        self.prev_btn.pack(side=tk.LEFT)

        self.page_info = tk.Label(self.nav_frame, text="Página 1/1")
        self.page_info.pack(side=tk.LEFT, padx=5)

        self.next_btn = tk.Button(
            self.nav_frame, text="Próxima >", command=lambda: self.change_page(1)
        )
        self.next_btn.pack(side=tk.LEFT)

        # Canvas com scrollbar para ordens
        self.orders_canvas = tk.Canvas(self.frame_orders_manager)
        self.orders_scrollbar = tk.Scrollbar(
            self.frame_orders_manager,
            orient=tk.VERTICAL,
            command=self.orders_canvas.yview,
        )
        self.orders_canvas.configure(yscrollcommand=self.orders_scrollbar.set)

        self.orders_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.orders_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Frame interno para os Checkbuttons
        self.orders_inner_frame = tk.Frame(self.orders_canvas)
        self.orders_canvas.create_window(
            (0, 0), window=self.orders_inner_frame, anchor="nw"
        )

        # Vincular evento de redimensionamento do canvas
        self.orders_inner_frame.bind(
            "<Configure>",
            lambda e: self.orders_canvas.configure(
                scrollregion=self.orders_canvas.bbox("all")
            ),
        )

        # Scroll com mouse wheel
        self.orders_canvas.bind(
            "<Enter>",
            lambda e: self.orders_canvas.bind_all("<MouseWheel>", self._on_mousewheel),
        )
        self.orders_canvas.bind(
            "<Leave>", lambda e: self.orders_canvas.unbind_all("<MouseWheel>")
        )

        self.order_vars = {}  # Dicionário para armazenar variáveis dos checkboxes
        self.order_checkbuttons = {}  # Dicionário para armazenar os Checkbuttons

        # Frame para botões logo abaixo do canvas
        self.btn_frame = tk.Frame(self.frame_orders_manager)
        self.btn_frame.pack(fill=tk.X, pady=5)

        self.btn_remove_orders = tk.Button(
            self.btn_frame,
            text="Remover Ordens Selecionadas",
            command=self.remove_selected_orders,
        )
        self.btn_remove_orders.pack(side=tk.LEFT, padx=5)

        self.btn_show_orders = tk.Button(
            self.btn_frame,
            text="Mostrar Ensaios Selecionados",
            command=self.show_selected_orders,
        )
        self.btn_show_orders.pack(side=tk.LEFT, padx=5)

        # Seção para visualização das ordens
        self.frame_view = tk.Frame(self.root)
        self.frame_view.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.label_view = tk.Label(
            self.frame_view, text="Visualizar Ordens Adicionadas com Filtros:"
        )
        self.label_view.pack()

        # Opções de filtro
        self.filter_frame = tk.Frame(self.frame_view)
        self.filter_frame.pack(pady=5)

        self.filter_version_label = tk.Label(self.filter_frame, text="Versão:")
        self.filter_version_label.pack(side=tk.LEFT, padx=5)

        self.filter_version_entry = tk.Entry(self.filter_frame, width=10)
        self.filter_version_entry.pack(side=tk.LEFT, padx=5)

        self.filter_order_label = tk.Label(self.filter_frame, text="Ordem:")
        self.filter_order_label.pack(side=tk.LEFT, padx=5)

        self.filter_order_entry = tk.Entry(self.filter_frame, width=10)
        self.filter_order_entry.pack(side=tk.LEFT, padx=5)

        self.filter_temperature_label = tk.Label(self.filter_frame, text="Temperatura:")
        self.filter_temperature_label.pack(side=tk.LEFT, padx=5)

        self.filter_temperature_entry = tk.Entry(self.filter_frame, width=10)
        self.filter_temperature_entry.pack(side=tk.LEFT, padx=5)

        self.btn_apply_filters = tk.Button(
            self.filter_frame, text="Aplicar Filtros", command=self.apply_filters
        )
        self.btn_apply_filters.pack(side=tk.LEFT, padx=5)

        # Lista de resultados com barras de rolagem
        self.frame_results = tk.Frame(self.frame_view)
        self.frame_results.pack(pady=10, fill=tk.BOTH, expand=True)

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
        self.frame_results.grid_rowconfigure(0, weight=1)
        self.frame_results.grid_columnconfigure(0, weight=1)

        self.btn_refresh = tk.Button(
            self.frame_view, text="Atualizar Visualização", command=self.refresh_view
        )
        self.btn_refresh.pack(pady=5)

        # Arquivo JSON e pasta de Excel
        self.json_file = "Data.json"
        self.excel_folder = r"C:\Users\rickl\Downloads\CPK"

        # Atualizar lista de ordens ao iniciar
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
            self.update_orders_list()  # Atualizar lista de ordens após processamento
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
        """Atualiza a lista de ordens no gerenciador, ordenada por test_date e paginada."""
        # Limpar Checkbuttons existentes
        for widget in self.orders_inner_frame.winfo_children():
            widget.destroy()
        self.order_vars.clear()
        self.order_checkbuttons.clear()

        if not os.path.exists(self.json_file):
            self.page_info.config(text="Página 0/0")
            return

        try:
            with open(self.json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Coletar ordens com versão e test_date
            self.orders_list = []
            for version, orders in data.items():
                for order, details in orders.items():
                    test_date = details["metadados"].get("test_date", "0000-00-00")
                    self.orders_list.append((version, order, test_date))

            # Ordenar por test_date
            self.orders_list.sort(key=lambda x: x[2] if x[2] else "0000-00-00")

            # Calcular total de páginas
            total_orders = len(self.orders_list)
            self.total_pages = max(
                1, (total_orders + self.orders_per_page - 1) // self.orders_per_page
            )

            # Ajustar página atual se necessário
            if self.current_page > self.total_pages:
                self.current_page = self.total_pages

            # Atualizar label da página
            self.page_info.config(text=f"Página {self.current_page}/{self.total_pages}")

            # Habilitar/desabilitar botões de navegação
            self.prev_btn.config(
                state=tk.NORMAL if self.current_page > 1 else tk.DISABLED
            )
            self.next_btn.config(
                state=tk.NORMAL if self.current_page < self.total_pages else tk.DISABLED
            )

            # Obter ordens da página atual
            start_idx = (self.current_page - 1) * self.orders_per_page
            end_idx = start_idx + self.orders_per_page
            paginated_orders = self.orders_list[start_idx:end_idx]

            # Adicionar Checkbuttons para cada ordem da página
            for idx, (version, order, test_date) in enumerate(paginated_orders):
                var = tk.BooleanVar()
                self.order_vars[(version, order)] = var
                display_text = f"Versão: {version}, Ordem: {order}, Data: {test_date}"
                chk = tk.Checkbutton(
                    self.orders_inner_frame,
                    text=display_text,
                    variable=var,
                    anchor="w",
                    width=50,
                )
                chk.grid(row=idx, column=0, sticky="w", padx=5, pady=2)
                self.order_checkbuttons[(version, order)] = chk

            # Resetar checkbox selecionar tudo
            self.select_all_var.set(False)

            # Atualizar região de rolagem
            self.orders_canvas.configure(scrollregion=self.orders_canvas.bbox("all"))

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar ordens: {str(e)}")

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

    def remove_selected_orders(self):
        """Remove as ordens selecionadas do banco de dados."""
        if not os.path.exists(self.json_file):
            messagebox.showerror("Erro", "Nenhum banco de dados encontrado.")
            return
        try:
            with open(self.json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            selected_orders = [
                (version, order)
                for (version, order), var in self.order_vars.items()
                if var.get()
            ]

            if not selected_orders:
                messagebox.showwarning("Aviso", "Nenhuma ordem selecionada.")
                return

            # Remover ordens selecionadas
            for version, order in selected_orders:
                if version in data and order in data[version]:
                    del data[version][order]
                    if not data[version]:  # Se a versão ficou vazia, removê-la
                        del data[version]

            # Salvar JSON atualizado
            with open(self.json_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            messagebox.showinfo("Sucesso", "Ordens selecionadas removidas com sucesso!")
            self.update_orders_list()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao remover ordens: {str(e)}")

    def show_selected_orders(self):
        """Mostrar detalhes dos ensaios selecionados na lista de visualização."""
        # Implementar conforme necessidade, exemplo simples:
        selected_orders = [
            (version, order)
            for (version, order), var in self.order_vars.items()
            if var.get()
        ]
        if not selected_orders:
            messagebox.showwarning("Aviso", "Nenhuma ordem selecionada para mostrar.")
            return

        try:
            with open(self.json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.list_results.delete(0, tk.END)
            for version, order in selected_orders:
                if version in data and order in data[version]:
                    details = data[version][order]
                    line = f"Versão: {version} | Ordem: {order} | Data: {details['metadados'].get('test_date', '')}"
                    self.list_results.insert(tk.END, line)
                    # Pode adicionar mais detalhes conforme necessário

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao mostrar ordens: {str(e)}")

    def refresh_view(self):
        try:
            if not os.path.exists(self.json_file):
                messagebox.showerror("Erro", "Nenhum banco de dados encontrado.")
                return

            with open(self.json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.list_results.delete(0, tk.END)

            version_filter = self.filter_version_entry.get().strip()
            order_filter = self.filter_order_entry.get().strip()
            temp_filter = self.filter_temperature_entry.get().strip()

            for version, orders in data.items():
                if version_filter and version != version_filter:
                    continue

                for order, details in orders.items():
                    if order_filter and order != order_filter:
                        continue

                    for temp_type, temp_data in details["temperaturas"].items():
                        if temp_filter and temp_type != temp_filter:
                            continue

                        self.list_results.insert(
                            tk.END,
                            f"Versão: {version}, Ordem: {order}, Temperatura: {temp_type}, "
                            f"Meta: {details['metadados']}, Temperatura_C: {temp_data.get('temperatura_c')}, "
                            f"Ensaios: {temp_data.get('ensaios')}, Dados_Pressao: {temp_data.get('dados_pressao')}, "
                            f"Limites: {temp_data.get('limites')}",
                        )

            # Sincronizar checkboxes com filtro de ordem
            self.sync_checkboxes_with_order_filter()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar visualização: {str(e)}")

    def apply_filters(self):
        try:
            if not os.path.exists(self.json_file):
                messagebox.showerror("Erro", "Nenhum banco de dados encontrado.")
                return

            with open(self.json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            version_filter = self.filter_version_entry.get().strip()
            order_filter = self.filter_order_entry.get().strip()
            temp_filter = self.filter_temperature_entry.get().strip()

            self.list_results.delete(0, tk.END)

            for version, orders in data.items():
                if version_filter and version != version_filter:
                    continue

                for order, details in orders.items():
                    if order_filter and order != order_filter:
                        continue

                    for temp_type, temp_data in details["temperaturas"].items():
                        if temp_filter and temp_type != temp_filter:
                            continue

                        self.list_results.insert(
                            tk.END,
                            f"Versão: {version}, Ordem: {order}, Temperatura: {temp_type}, "
                            f"Meta: {details['metadados']}, Temperatura_C: {temp_data.get('temperatura_c')}, "
                            f"Ensaios: {temp_data.get('ensaios')}, Dados_Pressao: {temp_data.get('dados_pressao')}, "
                            f"Limites: {temp_data.get('limites')}",
                        )

            # Sincronizar checkboxes com filtro de ordem
            self.sync_checkboxes_with_order_filter()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao aplicar filtros: {str(e)}")

    def sync_checkboxes_with_order_filter(self):
        """Sincroniza os checkboxes com o filtro de ordem."""
        order_filter = self.filter_order_entry.get().strip()
        for (version, order), var in self.order_vars.items():
            var.set(order == order_filter and order_filter != "")

    def clean_value(self, value):
        if isinstance(value, str):
            value = value.strip().lstrip("'")
            if value in ("", "None", "NaN"):
                return None
            return value
        return value

    def parse_date(self, date_value):
        if isinstance(date_value, datetime):
            return date_value.strftime("%Y-%m-%d")
        return date_value

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    converter = ExcelToJsonConverter()
    converter.run()
