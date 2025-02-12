#Code from yu-zhiqian@outlook.com, ZhiqianYu@github

import sys
import pandas as pd
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QComboBox, QFileDialog, QMessageBox

class ExcelMatcher(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.df_dict = {}  # 存储不同表格的 DataFrame

    def initUI(self):
        layout = QVBoxLayout()

        # 选择 Excel 文件
        self.label = QLabel("Choose Excel File:")
        layout.addWidget(self.label)

        self.btn_select_file = QPushButton("Select File")
        self.btn_select_file.clicked.connect(self.load_file)
        layout.addWidget(self.btn_select_file)

        # 选择配对表格和列
        self.combo_table = QComboBox()
        layout.addWidget(QLabel("Choose Table for Pair Check:"))
        layout.addWidget(self.combo_table)
        self.combo_table.currentTextChanged.connect(self.update_pair_columns)

        self.combo_pair_column = QComboBox()
        layout.addWidget(QLabel("Choose Column for Pair Check:"))
        layout.addWidget(self.combo_pair_column)

        # 选择搜索表格和列
        self.combo_search_table = QComboBox()
        layout.addWidget(QLabel("Choose Table to Search at:"))
        layout.addWidget(self.combo_search_table)
        self.combo_search_table.currentTextChanged.connect(self.update_search_columns)

        self.combo_search_column = QComboBox()
        layout.addWidget(QLabel("Choose Column to Search at:"))
        layout.addWidget(self.combo_search_column)

        self.combo_data_to_fill = QComboBox()
        layout.addWidget(QLabel("Choose Column of Data for Fill In:"))
        layout.addWidget(self.combo_data_to_fill)

        # 选择填充表格和列
        self.combo_fill_table = QComboBox()
        layout.addWidget(QLabel("Choose Table to Fill In:"))
        layout.addWidget(self.combo_fill_table)
        self.combo_fill_table.currentTextChanged.connect(self.update_fill_columns)

        self.combo_fill_column = QComboBox()
        layout.addWidget(QLabel("Choose Column to Fill In:"))
        layout.addWidget(self.combo_fill_column)

        # 选择匹配规则
        self.combo_match_type = QComboBox()
        self.combo_match_type.addItems(["Exact Match", "Contains", "Starts With", "Ends With"])
        layout.addWidget(QLabel("Choose Match Type:"))
        layout.addWidget(self.combo_match_type)

        # 执行匹配
        self.btn_match = QPushButton("Match & Fill")
        self.btn_match.clicked.connect(self.match_and_fill)
        layout.addWidget(self.btn_match)

        self.setLayout(layout)

    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return
        
        try:
            self.df_dict = pd.read_excel(file_path, sheet_name=None)
            self.combo_table.clear()
            self.combo_search_table.clear()
            self.combo_fill_table.clear()

            tables = list(self.df_dict.keys())
            self.combo_table.addItems(tables)
            self.combo_search_table.addItems(tables)
            self.combo_fill_table.addItems(tables)

            # Store the file path in the class instance variable
            self.original_file_path = file_path

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read Excel: {e}")

    def update_pair_columns(self):
        self.combo_pair_column.clear()
        table = self.combo_table.currentText()
        if table and table in self.df_dict:
            self.combo_pair_column.addItems(self.df_dict[table].columns.astype(str))

    def update_search_columns(self):
        self.combo_search_column.clear()
        self.combo_data_to_fill.clear()
        table = self.combo_search_table.currentText()
        if table and table in self.df_dict:
            columns = self.df_dict[table].columns.astype(str)
            self.combo_search_column.addItems(columns)
            self.combo_data_to_fill.addItems(columns)

    def update_fill_columns(self):
        self.combo_fill_column.clear()
        table = self.combo_fill_table.currentText()
        if table and table in self.df_dict:
            self.combo_fill_column.addItems(self.df_dict[table].columns.astype(str))

    def match_and_fill(self):
        pair_table = self.combo_table.currentText()
        pair_column = self.combo_pair_column.currentText()
        search_table = self.combo_search_table.currentText()
        search_column = self.combo_search_column.currentText()
        data_column = self.combo_data_to_fill.currentText()
        fill_table = self.combo_fill_table.currentText()
        fill_column = self.combo_fill_column.currentText()
        match_type = self.combo_match_type.currentText()

        if not all([pair_table, pair_column, search_table, search_column, data_column, fill_table, fill_column]):
            QMessageBox.warning(self, "Warning", "Please select all required options!")
            return

        df_pair = self.df_dict[pair_table]
        df_search = self.df_dict[search_table]
        df_fill = self.df_dict[fill_table]

        def match_rule(value, target):
            if pd.isna(value) or pd.isna(target):
                return False
            value, target = str(value), str(target)
            if match_type == "Exact Match":
                return value == target
            elif match_type == "Contains":
                return value in target
            elif match_type == "Starts With":
                return target.startswith(value)
            elif match_type == "Ends With":
                return target.endswith(value)
            return False

        match_dict = {}
        for _, row in df_search.iterrows():
            search_value = row[search_column]
            data_value = row[data_column]
            for _, pair_row in df_pair.iterrows():
                pair_value = pair_row[pair_column]
                if match_rule(pair_value, search_value):
                    match_dict[pair_value] = data_value
                    break

        df_fill[fill_column] = df_fill[pair_column].map(match_dict).fillna("Not Found")

        # Save to the original file if the original file path exists
        if hasattr(self, 'original_file_path') and self.original_file_path:
            with pd.ExcelWriter(self.original_file_path, engine="xlsxwriter") as writer:
                for sheet, df in self.df_dict.items():
                    df.to_excel(writer, sheet_name=sheet, index=False)
            QMessageBox.information(self, "Success", "Matching and filling completed successfully!")
        else:
            # If the original file path is not available, prompt for a save location
            save_path, _ = QFileDialog.getSaveFileName(self, "Save Updated Excel", "", "Excel Files (*.xlsx *.xls)")
            if save_path:
                with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
                    for sheet, df in self.df_dict.items():
                        df.to_excel(writer, sheet_name=sheet, index=False)
                QMessageBox.information(self, "Success", "Matching and filling completed successfully!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelMatcher()
    window.show()
    sys.exit(app.exec())
