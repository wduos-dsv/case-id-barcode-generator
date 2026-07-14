import sys
import os
import re
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, 
                             QFileDialog, QLabel, QMessageBox, QHBoxLayout, QComboBox)
from PyQt6.QtCore import Qt

# --- Import processing libraries ---
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
import barcode
from barcode.writer import ImageWriter

class BarcodeApp(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_file_path = ""
        self.initUI()
        
    def initUI(self):
        # Set up window properties
        self.setWindowTitle("Case ID Barcode Generator")
        self.setMinimumSize(480, 220)
        
        # Main layout
        layout = QVBoxLayout()
        layout.setSpacing(15)
        
        # File selector section
        file_layout = QHBoxLayout()
        self.file_label = QLabel("Nenhum arquivo selecionado...")
        self.file_label.setStyleSheet("color: #666; font-style: italic;")
        
        btn_browse = QPushButton("Procurar...")
        btn_browse.clicked.connect(self.open_file_dialog)
        
        file_layout.addWidget(self.file_label, stretch=4)
        file_layout.addWidget(btn_browse, stretch=1)
        layout.addLayout(file_layout)
        
        # Sorting dropdown section
        sort_layout = QHBoxLayout()
        sort_label = QLabel("Ordenar por:")
        sort_label.setStyleSheet("font-weight: bold;")
        
        self.sort_dropdown = QComboBox()
        self.sort_dropdown.addItems([
            "SKU",
            "Endereço",
            "Tipo de Caixa"
        ])
        
        sort_layout.addWidget(sort_label, stretch=1)
        sort_layout.addWidget(self.sort_dropdown, stretch=4)
        layout.addLayout(sort_layout)
        
        # Process Button
        self.btn_process = QPushButton("Iniciar Processamento")
        self.btn_process.setEnabled(False) 
        self.btn_process.setStyleSheet("""
            QPushButton { background-color: #0078d4; color: white; font-weight: bold; padding: 8px; border-radius: 4px; }
            QPushButton:disabled { background-color: #cccccc; color: #666666; }
            QPushButton:hover { background-color: #005a9e; }
        """)
        self.btn_process.clicked.connect(self.process_spreadsheet)
        layout.addWidget(self.btn_process)
        
        self.setLayout(layout)
        
    def open_file_dialog(self):
        file_filter = "Excel Files (*.xlsx *.xlsm)"
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha", "", file_filter)
        
        if file_path:
            self.selected_file_path = file_path
            display_name = os.path.basename(file_path)
            self.file_label.setText(f"Arquivo: {display_name}")
            self.file_label.setStyleSheet("color: #000; font-style: normal; font-weight: bold;")
            self.btn_process.setEnabled(True)

    def process_spreadsheet(self):
        if not self.selected_file_path:
            return
            
        try:
            # 1. Load the selected workbook
            wb = openpyxl.load_workbook(self.selected_file_path, data_only=True)
            if 'Imprimir' not in wb.sheetnames:
                raise ValueError("A aba 'Imprimir' não foi encontrada no arquivo selecionado.")
                
            ws = wb['Imprimir']
            
            # 2. Setup Barcode options
            options = {
                'module_width': 0.2,
                'module_height': 7.0,
                'quiet_zone': 2.0,
                'write_text': False 
            }
            
            BARCODE_WIDTH_PX = 130
            BARCODE_HEIGHT_PX = 40
            ROW_HEIGHT_PT = 45
            COL_WIDTH_CHAR = 22
            
            CELL_WIDTH_PX = COL_WIDTH_CHAR * 7.1 
            CELL_HEIGHT_PX = ROW_HEIGHT_PT * 1.333
            
            # --- INITIALIZE REMOVAL COUNTER ---
            full_pallets_removed = 0
            
            # 3. Extract, filter, and sort data
            rows_data = []
            for row in range(4, 102):
                row_values = [ws.cell(row=row, column=col).value for col in range(1, 7)]
                if row_values[1] is not None:
                    qty_value = row_values[2]
                    
                    # Rule 2: Keep only rows that do NOT equal 360 or 400
                    if qty_value not in (360, 400, "360", "400"):
                        rows_data.append(row_values)
                    else:
                        # Increment the counter every time we skip a row matching 360 or 400
                        full_pallets_removed += 1
            
            # Get selected sorting method from dropdown index
            sort_method = self.sort_dropdown.currentIndex()
            
            if sort_method == 0:
                # Option 1: By Item Code (Column B, numeric index 1)
                def get_item_key(r):
                    try:
                        return int(r[1])
                    except (ValueError, TypeError):
                        return 999999
                rows_data.sort(key=get_item_key)
                
            elif sort_method == 1:
                # Option 2: By Location/Street (Column A, string index 0)
                def get_location_key(r):
                    val = str(r[0]) if r[0] is not None else ""
                    parts = re.split(r'(\d+)', val)
                    return [int(text) if text.isdigit() else text.lower() for text in parts]
                rows_data.sort(key=get_location_key)
                
            elif sort_method == 2:
                # Option 3: By Box Type (Column D, string index 3)
                def get_box_key(r):
                    val = str(r[3]) if r[3] is not None else ""
                    return val.lower()
                rows_data.sort(key=get_box_key)
            
            # 4. Clear and rewrite grid
            for row in range(4, 102):
                for col in range(1, 7):
                    ws.cell(row=row, column=col).value = None
                    
            for index, row_values in enumerate(rows_data):
                current_row = 4 + index
                for col_idx, val in enumerate(row_values):
                    ws.cell(row=current_row, column=col_idx + 1).value = val
            
            # 5. Generate and center barcodes
            ws.column_dimensions['F'].width = COL_WIDTH_CHAR
            
            for index in range(len(rows_data)):
                row = 4 + index
                cell_value = ws[f'E{row}'].value
                
                if cell_value and not str(cell_value).startswith('='):
                    code128 = barcode.get('code128', str(cell_value), writer=ImageWriter())
                    filename = f"temp_barcode_{row}"
                    code128.save(filename, options=options)
                    
                    img_path = f"{filename}.png"
                    img = Image(img_path)
                    
                    ws.row_dimensions[row].height = ROW_HEIGHT_PT
                    
                    left_margin_px = max(0, (CELL_WIDTH_PX - BARCODE_WIDTH_PX) / 2)
                    top_margin_px = max(0, (CELL_HEIGHT_PX - BARCODE_HEIGHT_PX) / 2)
                    left_margin_emu = int(left_margin_px * 9525)
                    top_margin_emu = int(top_margin_px * 9525)
                    
                    marker = AnchorMarker(col=5, colOff=left_margin_emu, row=row-1, rowOff=top_margin_emu)
                    barcode_size = XDRPositiveSize2D(
                        cx=pixels_to_EMU(BARCODE_WIDTH_PX), 
                        cy=pixels_to_EMU(BARCODE_HEIGHT_PX)
                    )
                    
                    img.anchor = OneCellAnchor(_from=marker, ext=barcode_size)
                    ws.add_image(img)
            
            # 6. Remove other sheets
            for sheet_name in wb.sheetnames:
                if sheet_name != 'Imprimir':
                    wb.remove(wb[sheet_name])
            
            # 7. Save output file
            output_dir = os.path.dirname(self.selected_file_path)
            output_path = os.path.join(output_dir, 'Imprimir.xlsx')
            wb.save(output_path)
            
            # Clean up temp image assets
            for index in range(len(rows_data)):
                row = 4 + index
                if os.path.exists(f"temp_barcode_{row}.png"):
                    os.remove(f"temp_barcode_{row}.png")
                    
            # --- MODIFIED SUCCESS POPUP WITH THE REQUESTED COUNTER ---
            success_msg = (
                f"Processamento concluído com sucesso!\n"
                f"Salvo em: {output_path}\n\n"
                f"Número de pallets FULL neste pedido: {full_pallets_removed}"
            )
            QMessageBox.information(self, "Sucesso", success_msg)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro durante o processamento:\n{str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = BarcodeApp()
    ex.show()
    sys.exit(app.exec())