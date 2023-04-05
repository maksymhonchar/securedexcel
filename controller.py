class SpreadsheetController(object):

    def __init__(self, model, view):
        self._model = model
        self._view = view
        self.connect_signals()

    def connect_signals(self):
        self._view.open_file_action.triggered.connect(self.import_spreadsheet)
        self._view.save_file_action.triggered.connect(self.export_spreadsheet)
        self._view.add_column_action.triggered.connect(self.add_column)
        self._view.search_edit.returnPressed.connect(self.search_data)
        self._view.search_button.clicked.connect(self.search_data)

    def connect_tab_widget_cellChanged(self):
        for tab_idx in range(self._view.tab_widget.count()):
            self._view.tab_widget.widget(tab_idx).cellChanged.connect(
                self.handle_cell_changed
            )

    def import_spreadsheet(self):
        file_path = self._view.get_open_file_path()
        if file_path:
            # clear existing tabs
            self._view.tab_widget.clear()
            # load excel file content
            self._model.load_spreadsheet(file_path)
            # load sheet content into tabs
            for sheet_title in self._model.get_worksheets_titles():
                worksheet = self._model.get_worksheet(sheet_title)
                self._view.add_tab(worksheet)
            # connect each table.cellChanged signal
            self.connect_tab_widget_cellChanged()

    def export_spreadsheet(self):
        file_path = self._view.get_save_file_path()
        if file_path:
            if not file_path.endswith(".xlsx"):
                file_path += ".xlsx"
            self._model.save_spreadsheet(file_path)

    def add_column(self):
        current_tab_text = self._view.get_current_tab_text()
        self._model.add_column(current_tab_text)
        self._view.add_column_in_current_tab()

    def search_data(self):
        search_query = self._view.get_current_search_query()
        search_results = self._model.search(search_query)
        if search_results:
            search_results_text = ''
            for sheet_name, sheet_search_results in search_results.items():
                if sheet_search_results:
                    search_results_text += f'Сторінка <b>{sheet_name}</b>:'
                    search_results_text += '<ul>'
                    search_results_text += ' '.join(
                        [
                            f'<li>{search_result}</li>'
                            for search_result in sheet_search_results
                        ]
                    )
                    search_results_text += '</ul>'
                    search_results_text += '<br>'
        else:
            search_results_text = 'Не знайдено жодного результату'
        self._view.set_search_results(search_results_text)

    def handle_cell_changed(self, row_idx: int, col_idx: int):
        current_tab_text = self._view.get_current_tab_text()
        new_value = self._view.get_cell_value_in_current_tab(row_idx, col_idx)
        self._model.update_cell(current_tab_text, row_idx, col_idx, new_value)
