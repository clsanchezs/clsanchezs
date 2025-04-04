import sys
import re
from fuzzywuzzy import process
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QComboBox, QStackedWidget, 
    QGroupBox, QFormLayout, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QHBoxLayout, QSpacerItem, QSizePolicy, QFileDialog
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

class MyGISDataTool(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        self.stacked_widget = QStackedWidget()

        # --- Stage 1: Lobby ---
        self.lobby_page = QWidget()
        self.lobby_layout = QVBoxLayout()

        # Logos at the top corners
        logo_layout = QHBoxLayout()
        self.logo_left = QLabel()
        self.logo_right = QLabel()
        self.logo_left.setPixmap(QPixmap("Eni_2023.svg.png").scaled(200, 200, Qt.KeepAspectRatio))
        self.logo_right.setPixmap(QPixmap("HYDROM_LOGO.png").scaled(200, 200, Qt.KeepAspectRatio))
        
        logo_layout.addWidget(self.logo_left, alignment=Qt.AlignLeft)
        logo_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        logo_layout.addWidget(self.logo_right, alignment=Qt.AlignRight)
        
        self.lobby_layout.addLayout(logo_layout)

        # Welcome message
        welcome_label = QLabel("""
        <h2>Benvenuto in MyGIS Data Validation Tool</h2>
        <p>Questa applicazione ti aiuterà a elaborare i tuoi dati in modo semplice ed efficace.</p>
        <p>Seleziona il Sito e il tipo di file per iniziare.</p>
        """)
        welcome_label.setAlignment(Qt.AlignCenter)
        self.lobby_layout.addWidget(welcome_label)

        # "Sito" Filter
        self.sito_label = QLabel("Sito:")
        self.sito_combo = QComboBox()
        self.sito_combo.addItems(["Seleziona", "ASSEMINI", "AVENZA", "AVENZA_EX_ITA_COKE", "BARI", "BRINDISI",
                                   "BRONTE_CG","CENGIO", "CESANO MADERNO", "CIRO' MARINA", "CROTONE", "ENNA_CG_GAGLIANO",
                                   "FERRANDINA","FERRARA", "FOGGIA", "FORNOVO", "GAVORRANO", "GELA", "GELA_16", "GELA_22",
                                   "GELA_32_103","GELA_35_67", "GELA_39_91_93_94", "GELA_4_98_101", "GELA_62", "GELA_63",
                                   "GELA_8","GELA_AREA_2_CRO", "GELA_AREA_I_CRO", "GELA_CAMMARATA1", "GELA_CRC_148",
                                   "GELA_FLU2_CM2_CM1","GELA_FLU3_ARMA2_CM1", "GELA_GC10", "GELA_GC51", "GELA_GIAURONE4",
                                   "GELA_NCO","GELA_SOTTOCLUSTER_D", "MANFREDONIA", "MANTOVA", "MANTOVA EDISON",
                                   "MISP_BacinoSR14","MISP_CPL", "MISP_CTE", "MISP_Discarica_ex_ausidet", 
                                   "MISP_Discarica_malcontenta","MISP_IMP", "MISP_Isola45_48", "MISP_Malcontenta_C", 
                                   "NAPOLI DECO", "NOVARA", "OTTANA","PADERNO DUGNANO", "PIEVE VERGONTE", 
                                   "PONTE GALERIA", "PORTO MARGHERA", "PORTO TORRES","PRIOLO", "RAGUSA14", "RAGUSA16", 
                                   "RAGUSA18", "RAVENNA", "RAVENNA_EX_DISTR", "RHO","ROBASSOMERO", "ROMA OSTIENSE", "SA CANNA", 
                                   "SA PIRAMIDE", "SAN GAVINO MONREALE", "SANNAZZARO", "SANTA GILLA", "SARROCH", "TROINA_ROVETTO1", "VIGGIANO", "VILLADOSSOLA"])
        
        self.lobby_layout.addWidget(self.sito_label)
        self.lobby_layout.addWidget(self.sito_combo)

        # Tipo di file loading selection
        self.file_loading_label = QLabel("Tipo di caricamento file:")
        self.file_loading_combo = QComboBox()
        self.file_loading_combo.addItems(["Seleziona", "Campagna", "Massivo"])
        self.file_loading_combo.currentIndexChanged.connect(self.update_file_type_options)
        
        self.lobby_layout.addWidget(self.file_loading_label)
        self.lobby_layout.addWidget(self.file_loading_combo)

        # Tipo di file selection
        self.file_type_label = QLabel("Tipo di file:")
        self.file_type_combo = QComboBox()
        self.file_type_combo.hide()
        self.file_type_combo.currentIndexChanged.connect(self.update_sub_options)
        
        self.lobby_layout.addWidget(self.file_type_label)
        self.lobby_layout.addWidget(self.file_type_combo)

        # Sub-options (initially hidden)
        self.sub_options_label = QLabel("Formato dei dati:")
        self.sub_options_combo = QComboBox()
        self.sub_options_combo.hide()
        self.sub_options_label.hide()
        
        self.lobby_layout.addWidget(self.sub_options_label)
        self.lobby_layout.addWidget(self.sub_options_combo)

        # 'Avanti' Button
        self.next_button = QPushButton("Avanti")
        self.next_button.clicked.connect(self.go_to_data_verification)
        self.lobby_layout.addWidget(self.next_button)
        
        self.lobby_page.setLayout(self.lobby_layout)
        self.stacked_widget.addWidget(self.lobby_page)

        # --- Stage 2: Data Verification ---
        self.data_page = QWidget()
        self.data_layout = QVBoxLayout()

        # File Selection
        self.file1_label = QLabel("File 1: Not Selected")
        self.file1_button = QPushButton("Load Verification File")
        self.file1_button.clicked.connect(self.load_file1)


        self.file2_label = QLabel("File 2: Not Selected")
        self.file2_button = QPushButton("Load Analyte Alias File")
        self.file2_button.clicked.connect(self.load_file2)

        self.data_layout.addWidget(self.file1_label)
        self.data_layout.addWidget(self.file1_button)
        self.data_layout.addWidget(self.file2_label)
        self.data_layout.addWidget(self.file2_button)

        # Parameter Inputs
        self.params_group = QGroupBox("Set Parameters")
        form_layout = QFormLayout()

        self.param_inputs = {}

        # Group 1: Obbligatorio per la verifica
        group_verifica = QGroupBox("Obbligatorio per la verifica")
        verifica_layout = QFormLayout()

        verifica_params = [
            ("Row Start for Analytes", "row_start_analytes"),  # Mandatory
            ("Column for Analyte Names", "name_col_analytes"),  # Mandatory
            ("Column for Analyte Units", "um_col_analytes"),  # Mandatory
            ("Column Start for PDCs", "col_start_pdc"),  # Mandatory
            ("Row Start for PDCs", "row_start_pdc")  # Mandatory
            ]

        for label, key in verifica_params:
            self.param_inputs[key] = QLineEdit()
            verifica_layout.addRow(label, self.param_inputs[key])
        
        group_verifica.setLayout(verifica_layout)

        # Group 2: Obbligatorio per template MyGIS
        group_template = QGroupBox("Obbligatorio per template MyGIS")
        template_layout = QFormLayout()

        template_params = [
            ("Row Start for Dates", "row_start_dates"),  # Mandatory for template
            ("Row Start for Committente", "row_start_committente")  # Mandatory for template
            ]
        
        for label, key in template_params:
            self.param_inputs[key] = QLineEdit()
            template_layout.addRow(label, self.param_inputs[key])

        group_template.setLayout(template_layout)

        # Group 3: Opzionale
        group_opzionale = QGroupBox("Opzionale")
        opzionale_layout = QFormLayout()

        opzionale_params = [
            ("Row Start for RDPs", "row_start_rdp"),  # Optional
            ("Row Start for Codice Campione", "row_start_code_campione"),  # Optional
            ("Row Start for Tipo Campione", "row_start_tipo_campione"),  # Optional
            ("Row Start for Profundita Sup", "row_start_profundita_sup"),  # Optional
            ("Row Start for Profundita Inf", "row_start_profundita_inf"),  # Optional
            ("Row Start for Profundita Campione", "row_start_profundita_camp"),  # Optional
            ("Row Start for Tipo Allegato", "row_start_tipo_allegato"),  # Optional
            ("Row Start for Percorso Allegato", "row_percorso_allegato"),  # Optional
            ("Row Start for Note Campione", "row_note_campione")  # Optional
            ]
        
        for label, key in opzionale_params:
            self.param_inputs[key] = QLineEdit()
            opzionale_layout.addRow(label, self.param_inputs[key])

        group_opzionale.setLayout(opzionale_layout)

        # Create a horizontal layout for the groups
        groups_layout = QHBoxLayout()
        groups_layout.addWidget(group_verifica)
        groups_layout.addWidget(group_template)
        groups_layout.addWidget(group_opzionale)

        # Add the horizontal layout to the main layout
        self.data_layout.addLayout(groups_layout)

        # Process Button
        self.process_button = QPushButton("Process Data")
        self.process_button.clicked.connect(self.process_data)
        self.data_layout.addWidget(self.process_button)

        # Add "Back to Lobby" Button
        self.back_button = QPushButton("Back to Lobby")
        self.back_button.clicked.connect(self.go_to_lobby)  # Connect to the method to switch pages
        self.data_layout.addWidget(self.back_button)

        # Tables for Results
        self.result_table = QTableWidget()  # First table
        self.result_table_2 = QTableWidget()  # Second table
        self.result_table_3 = QTableWidget()  # Third table

        # Horizontal layout for the tables
        table_layout = QHBoxLayout()
        table_layout.addWidget(self.result_table)
        table_layout.addWidget(self.result_table_2)
        table_layout.addWidget(self.result_table_3)

        self.data_layout.addLayout(table_layout)  # Add the table layout to the second page

        self.data_page.setLayout(self.data_layout)
        self.stacked_widget.addWidget(self.data_page)

        layout.addWidget(self.stacked_widget)
        self.setLayout(layout)
        self.setWindowTitle("MyGIS Data Tool")
        self.setGeometry(100, 100, 1400, 900)

    def go_to_lobby(self):
        """Switch to the lobby page."""
        self.stacked_widget.setCurrentIndex(0) # Index 0 corresponds to the lobby_page

    def load_file1(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Verification File", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.file1_label.setText(f"File 1: {file_name}")
            self.file1_path = file_name

    def load_file2(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Analyte Alias File", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.file2_label.setText(f"File 2: {file_name}")
            self.file2_path = file_name 

    def update_file_type_options(self):
        selected_loading_type = self.file_loading_combo.currentText()
        
        if selected_loading_type in ["Campagna", "Massivo"]:
            self.file_type_combo.clear()
            self.file_type_combo.addItems(["Seleziona", "Campi incrociati", "Campi tabellare"])
            self.file_type_label.show()
            self.file_type_combo.show()
        else:
            self.file_type_combo.hide()
            self.file_type_label.hide()
            self.sub_options_combo.hide()
            self.sub_options_label.hide()

    def update_sub_options(self):
        selected_type = self.file_type_combo.currentText()
        
        if selected_type == "Campi incrociati":
            self.sub_options_combo.clear()
            self.sub_options_combo.addItems([
                "Seleziona", "Analiti in colonna - Pdc in righe", "PdC in colonna - Analiti in righe"
            ])
            self.sub_options_label.show()
            self.sub_options_combo.show()
        else:
            self.sub_options_combo.hide()
            self.sub_options_label.hide()

    def go_to_data_verification(self):
        self.stacked_widget.setCurrentIndex(1)

    def process_data(self):
        try:
            if not hasattr(self, 'file1_path') or not hasattr(self, 'file2_path'):
                self.result_table.setRowCount(1)
                self.result_table.setColumnCount(1)
                self.result_table.setItem(0, 0, QTableWidgetItem("Please load both files."))
                return
            
            df = pd.read_excel(self.file1_path, header=None)
            analiti_df = pd.read_excel(self.file2_path)
            
            row_start_analytes = int(self.param_inputs['row_start_analytes'].text() or 6) # Mandatory
            name_col_analytes = int(self.param_inputs['name_col_analytes'].text() or 0) # Mandatory
            um_col_analytes = int(self.param_inputs['um_col_analytes'].text() or 1) # Mandatory
            col_start_pdc = int(self.param_inputs['col_start_pdc'].text() or 2) # Mandatory
            row_start_pdc = int(self.param_inputs['row_start_pdc'].text() or 0)  # Mandatory

            # Dynamically construct the DataFrame
            dynamic_data = {}

            # Add mandatory PDC column
            dynamic_data["PDC"] = df.iloc[row_start_pdc, col_start_pdc:].dropna().values

            # Map user inputs to row names, optional variables
            row_mapping = {
                "Date": "row_start_dates",
                "RDP": "row_start_rdp",
                "Code Campione": "row_start_code_campione",
                "Committente": "row_start_committente",
                "Tipo Campione": "row_start_tipo_campione",
                "Profundita Sup": "row_start_profundita_sup",
                "Profundita Inf": "row_start_profundita_inf",
                "Profundita Camp": "row_start_profundita_camp",
                "Tipo Allegato": "row_start_tipo_allegato",
                "Percorso Allegato": "row_percorso_allegato",
                "Note Campione": "row_note_campione"
                }
            

            # Iterate over the mapping and extract data for each row
            for column_name, row_key in row_mapping.items():
                row_start = self.param_inputs[row_key].text()
                if row_start:  # Only include rows where the user has provided input
                    row_start = int(row_start)
                    dynamic_data[column_name] = df.iloc[row_start, col_start_pdc:].dropna().values


            # Create the DataFrame dynamically
            pdc_date_df = pd.DataFrame(dynamic_data)
            pdc_date_df = pdc_date_df.reset_index(drop=True)

            # Extract analytes
            analytes_df = df.iloc[row_start_analytes:, [name_col_analytes, um_col_analytes]].dropna()
            analytes_df.columns = ["Name", "Unit"]
            analytes_df.reset_index(drop=True, inplace=True)
            
            # Identify the start row of measurements (right after analytes)
            row_start_measurements = row_start_analytes  # Measurements start from analytes row

            # Extract the measurement matrix
            measurements = df.iloc[row_start_measurements:, col_start_pdc:].reset_index(drop=True)


            # Rename columns of the measurement matrix to ensure they are integers
            measurements.columns = range(len(measurements.columns))  # Renaming columns from 0 to n-1


            # Convert the measurement matrix to a long format
            long_df = measurements.melt(ignore_index=False)
            long_df.columns = ["PDC_Index", "Value"]

            # Assign actual analyte names and units
            long_df["Analyte"] = long_df.index.map(lambda i: analytes_df.iloc[i % len(analytes_df), 0])
            long_df["Unit"] = long_df.index.map(lambda i: analytes_df.iloc[i % len(analytes_df), 1])

            # Dynamically assign values to long_df based on pdc_date_df columns
            for column_name in pdc_date_df.columns:
                long_df[column_name] = long_df["PDC_Index"].map(
                    lambda i: pdc_date_df[column_name].iloc[i % len(pdc_date_df)]
                )

            # Drop the temporary index column
            long_df = long_df.drop(columns=["PDC_Index"]).reset_index(drop=True)

            # Step 1: Create an alias dictionary (flattening alias columns)
            alias_dict = {}

            for _, row in analiti_df.iterrows():
                official_name = row["ANALITA_NOME"].strip().lower()
                alias_dict[official_name] = official_name  # Direct lookup for official name

                # Loop over all ALIAS columns dynamically
                for col in analiti_df.columns:
                    if "ALIAS" in col:  # Ensure we only process alias columns
                        alias = row[col]
                        if pd.notna(alias) and alias.strip():
                            alias_dict[alias.strip().lower()] = official_name  # Map alias to official name

            def match_contaminant_with_type(name, alias_dict, similarity_threshold=90):
                # Normalize input
                name_cleaned = name.lower().strip()

                # Exact match 
                if name_cleaned in alias_dict:
                    return alias_dict[name_cleaned], "Exact"
                
                # Fuzzy Matching (if no exact match)
                result = process.extractOne(name_cleaned, alias_dict.keys())  
                if result:  # Check if None
                    match, score = result[:2]  # Ensure only two values are unpacked
                    if score >= similarity_threshold:
                        return alias_dict[match], "Fuzzy"
                return "Sconosciuto", "No Match"
            
            # Apply function and create the new DataFrame
            summary_df = analytes_df.copy()

            # Apply function
            summary_df[["Nome Ufficiale", "Matching Type"]] = summary_df["Name"].apply(lambda x: pd.Series(match_contaminant_with_type(x, alias_dict)))

            # Display the results in the three tables
            self.display_results(summary_df, self.result_table)  # First table
            self.display_results(pdc_date_df, self.result_table_2)  # Second table
            self.display_results(long_df, self.result_table_3)  # Third table

        except Exception as e:
            self.result_table.setRowCount(1)
            self.result_table.setColumnCount(1)
            self.result_table.setItem(0, 0, QTableWidgetItem(f"Error: {str(e)}"))

    def display_results(self, df, table_widget):
        """Display the DataFrame in the specified QTableWidget."""
        table_widget.setRowCount(len(df))
        table_widget.setColumnCount(len(df.columns))
        table_widget.setHorizontalHeaderLabels(df.columns)

        for i in range(len(df)):
            for j, col in enumerate(df.columns):
                table_widget.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyGISDataTool()
    window.show()
    sys.exit(app.exec_())
