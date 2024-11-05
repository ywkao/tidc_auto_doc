import os
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class QualityControlDocGenerator:
    def __init__(self, target_folder, filename, prefix='/content/drive/My Drive/'):
        self.base = prefix + target_folder
        self.filename = filename
        self.csv = os.path.join(target_folder, filename)
        self.doc = None
        self.black = RGBColor(0, 0, 0)
        self.blue = RGBColor(0, 0, 255)
        self.info = [
            ("User:", "TIDC"),
            ("Date:", "2024/09/13"),
            ("Version:", "V3-HD-Full-HB-v2.2"),
            ("Manufacturer:", "Plotech"),
            ("Batch:", "2432"),
            ("ID:", "320-07-01-003-00000")
        ]
        self.inspection_items = [
            ("General comments:", "", 2, [34, 0, 42]),
            ("Flatness:", "<0.5mm", 1, [2, 2]),
            ("Comments:", "", 0, [25, 0]),
            ("Thickness measurements:", "1.311mm", 1, [4, 22]),
            ("Plating (BGA):", "PASS", 1, [4, 8]),
            ("Plating (Holes):", "PASS", 0, [4, 8]),
            ("Soldermask alignment:", "PASS", 1, [4, 26]),
            ("Glue problems?", "PASS", 1, [7, 26]),
            ("Test coupons (observations, continuity measurements etc.):", "PASS", 2, [4, 10, 42]),
            ("Accept?", "PASS", 1, [4, 12])
        ]
        self.assembled_items = [
            ("General comments:", "", 1, [34, 0]),
            ("Flatness:", "", 1, [38, 0]),
            ("HGCROC type:", "", 1, [14, 0]),
            ("HGCROC rotation:", "", 0, [14, 0]),
            ("Connectors:", "", 1, [15, 0]),
            ("Resistors/capacitors:", "", 0, [12, 0])
        ]

        # Check & Read the CSV
        self._check_folder()
        self.df = self._read_and_process_csv()

    #----------------------------------------------------------------------------------------------------
    # main methods
    #----------------------------------------------------------------------------------------------------
    def create_directories(self):
        """ Create folders based on CERN ID """
        for index, row in self.df.iterrows():
            sub_folder = os.path.join(self.base, row['CERN ID'])
            os.makedirs(sub_folder, exist_ok=True)

    def create_documents(self):
        generator._convert_data()
        for i, element in enumerate(all_info_lists):
            generator._update_data(*element)
            generator._create_quality_control_doc()

    #----------------------------------------------------------------------------------------------------
    # auxiliary modules
    #----------------------------------------------------------------------------------------------------
    def _convert_data(self):
        # Find the Glue column
        glue_column = self._find_column_by_keyword('Glue')
        if glue_column: print(f"Found Glue column: {glue_column}")
        else: print("Warning: Glue column not found!")
    
        # Create inspection items for the first row
        self.all_info_lists = []
        for index, row in self.df.iterrows():
            info, items, gdoc = self._create_inspection_items(row)
            self.all_info_lists.append((info, items, gdoc))

    def _update_data(self, info, items, gdoc):
        self.info = info
        self.inspection_items = items
        self.cernID = info[5][1]
        self.folder = os.path.join(self.base, self.cernID)
        self.output_file = os.path.join(self.folder, f'{gdoc}')
        self.image_path = os.path.join(self.base, '20240913_104336.jpg')
        print(f"{self.image_path}")
        print(f"{self.output_file}")

    def _create_quality_control_doc(self):
        self.doc = docx.Document()
        self._set_page_margins()
        self._add_title()
        self._add_first_visual_inspection()
        self.doc.add_page_break()
        self._add_title()
        self._add_second_visual_inspection()
        self.doc.save(self.output_file)
        print(f"Document has been saved as {self.output_file}")

    #----------------------------------------------------------------------------------------------------
    # Load-data related
    #----------------------------------------------------------------------------------------------------
    def _check_folder(self):
        """ To-do: return error if folder/csv does not exist """

        # 確保目標資料夾存在
        if not os.path.exists(self.base):
            os.makedirs(self.base)
            print(f'新增 {self.base}')
        
        # List all files in the directory
        files = os.listdir(self.base)
        for f in files:
            full_path = os.path.join(self.base, f)
            file_size = os.path.getsize(full_path)
            print(f"- {f} ({file_size} bytes)")
        
        # Check for specific file
        if os.path.exists(self.csv):
            print(f"\nFound target file: {self.filename}")
            print(f"Full path: {self.csv}")
            print(f"File size: {os.path.getsize(self.csv)} bytes")

    def _read_and_process_csv(self):
        """
        Read and process the CSV file to extract relevant information
        """
        try:
            # Skip the first two rows and use the third row as headers
            df = pd.read_csv(self.csv, skiprows=2, header=0)
    
            # Clean up column names (remove newlines and extra spaces)
            df.columns = [str(col).strip().replace('\n', ' ') for col in df.columns]
    
            # Print the column names to verify
            print("\nColumn names:")
            for i, col in enumerate(df.columns):
                print(f"{i}: '{col}'")
    
            keywords = ['CERN ID', 'Accept?', 'Filename', 'Photo']
            # Print the first few rows of data
            print("\nFirst few rows of data (selected columns):")
            print(df[keywords].head())
    
            return df
    
        except Exception as e:
            print(f"Error processing file: {str(e)}")
            return None
    
    def _find_column_by_keyword(self, keyword):
        """
        Find column name that contains the given keyword
        """
        matching_cols = [col for col in self.df.columns if keyword.lower() in col.lower()]
        return matching_cols[0] if matching_cols else None
    
    def _create_inspection_items(self, row):
        """
        Create inspection items list based on DataFrame content for a specific row
        # row = df.iloc[row_index]
        """
    
        # Create lists with matching from DataFrame
        info = [
            ("User:", str(row['User']) if pd.notna(row['User']) else ""),
            ("Date:", str(row['Date']) if pd.notna(row['Date']) else ""),
            ("Version:", str(row['Version']) if pd.notna(row['Version']) else ""),
            ("Manufacturer:", str(row['Manufacturer']) if pd.notna(row['Manufacturer']) else ""),
            ("Batch:", str(row['Batch']) if pd.notna(row['Batch']) else ""),
            ("ID:", str(row['CERN ID']) if pd.notna(row['CERN ID']) else "")
        ]
    
        inspection_items = [
            ("General comments:", "", 2, [34, 0, 42]),
            ("Flatness:", f"{row['Flatness']}mm", 1, [2, 2]),
            ("Comments:", "", 0, [25, 0]),
            ("Thickness measurements:", f"{row['Thickness']}mm", 1, [4, 22]),
            ("Plating (BGA):", "PASS" if row['Plating (BGA)'] else "FAIL", 1, [4, 8]),
            ("Plating (Holes):", "PASS" if row['Plating (Holes)'] else "FAIL", 0, [4, 8]),
            ("Soldermask alignment:", "PASS" if row['Soldermask alignment'] else "FAIL", 1, [4, 26]),
            ("Glue problems?", "PASS" if row[glue_column] else "FAIL", 1, [7, 26]) if glue_column else
            ("Glue problems?", "UNKNOWN", 1, [7, 26]),
            ("Test coupons (observations, continuity measurements etc.):",
             "PASS" if row['Test coupons (observations, continuity measurements etc.)'] else "FAIL", 2, [4, 10, 42]),
            ("Accept?", "PASS" if row['Accept?'] else "FAIL", 1, [4, 12])
        ]
    
        gdoc = row['Filename'] + '.docx'
    
        return info, inspection_items, gdoc
    
    #----------------------------------------------------------------------------------------------------
    # Document-related
    #----------------------------------------------------------------------------------------------------
    def _set_page_margins(self):
        for section in self.doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

    def _add_underlined_spaces(self, paragraph, Nspaces, color=None):
        full_width_underscore = '＿'
        underlined_spaces = paragraph.add_run(full_width_underscore * Nspaces)
        underlined_spaces.font.color.rgb = color or self.black
        underlined_spaces.font.underline = WD_UNDERLINE.THICK

    def _add_title(self):
        title = self.doc.add_paragraph("Hexaboard 8\"V3 HD-FUll-HB-V2.2 Quality control traveler document V3")
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_format = title.runs[0].font
        title_format.size = Pt(14)
        title_format.bold = True

        for i, (key, value) in enumerate(self.info):
            if i % 3 == 0:
                p = self.doc.add_paragraph()

            run = p.add_run(key)
            run.font.color.rgb = self.black
            self._add_underlined_spaces(p, 2)

            run = p.add_run(value)
            run.font.underline = WD_UNDERLINE.THICK
            run.font.color.rgb = self.blue
            self._add_underlined_spaces(p, 4 - (i // 5))

    def _add_formatted_table(self):
        # 添加一個空行作為間隔
        self.doc.add_paragraph()

        # 添加新的1x2表格
        new_table = self.doc.add_table(rows=1, cols=2)
        new_table.style = 'Normal Table'

        for cell in new_table.rows[0].cells:
            # 添加5個空行到每個單元格
            for _ in range(5):
                cell.add_paragraph()

            # skip customized format
            continue

            # 設置單元格邊框為黑色
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()

            # 設置邊框
            for border_pos in ['top', 'left', 'bottom', 'right']:
                border = OxmlElement('w:tcBorders')
                border_element = OxmlElement(f'w:{border_pos}')
                border_element.set(qn('w:val'), 'single')
                border_element.set(qn('w:sz'), '4')
                border_element.set(qn('w:space'), '0')
                border_element.set(qn('w:color'), '#000000')
                border.append(border_element)
                tcPr.append(border)

    def _add_first_visual_inspection(self):
        self.doc.add_heading("1st Visual Inspection – Bare PCB", level=1)

        for item, value, nLines, spaces in self.inspection_items:
            if 'Accept' in item:
                paragraph = self.doc.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = paragraph.add_run()
                run.add_picture(self.image_path, width=Inches(6))
                print(f'[INFO] add picture: {self.image_path}')

            if nLines > 0:
                p = self.doc.add_paragraph()

            p.add_run(f"{item} ")
            self._add_underlined_spaces(p, spaces[0])

            run = p.add_run(value)
            run.font.underline = WD_UNDERLINE.THICK
            run.font.color.rgb = self.blue
            self._add_underlined_spaces(p, spaces[1])

            if nLines == 2:
                p = self.doc.add_paragraph()
                self._add_underlined_spaces(p, spaces[2])

    def _add_second_visual_inspection(self):
        self.doc.add_heading("2nd Visual Inspection – Assembled PCB", level=1)

        for item, value, nLines, spaces in self.assembled_items:
            if nLines > 0:
                p = self.doc.add_paragraph()
            p.add_run(f"{item} ")
            self._add_underlined_spaces(p, spaces[0])
            p.add_run(value)
            self._add_underlined_spaces(p, spaces[1])

        self._add_formatted_table()

        self.doc.add_heading("Functional Tests", level=1)

        functional_tests = [
            ("Power-on current:", "_______________"+' '*8),
            ("Configured OK:", "Yes/No"+' '*8),
            ("Operating current:", "________________"),
            ("DAQ lines OK:", ""),
        ]

        for i, (item, value) in enumerate(functional_tests):
            if i%3==0:
              p = self.doc.add_paragraph()
            p.add_run(f"{item} ")
            p.add_run(value)

# Usage example
if __name__ == "__main__":
    folder, filename = 'autoDoc', 'test.csv'
    generator = QualityControlDocGenerator(folder, filename)
    generator.create_directories()
    generator.create_documents()
