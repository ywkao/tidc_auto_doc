import os
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class QualityControlDocGenerator:
    def __init__(self, target_folder, filename, prefix='/content/drive/My Drive/'):
        self.prefix = prefix if prefix.endswith('/') else prefix + '/' # 確保 prefix 結尾有斜線
        self.base = os.path.join(self.prefix, target_folder)
        self.filename = filename
        self.csv = os.path.join(self.base, filename)
        self.doc = None
        self.black = RGBColor(0, 0, 0)
        self.blue = RGBColor(0, 0, 255)

        print(f">>> {self.base}")
        print(f">>> {self.csv}")

        # Check & Read the CSV
        self._check_folder()
        self.df = self._read_and_process_csv()

        # Find the Glue column
        self.glue_column = self._find_column_by_keyword('Glue')
        if not self.glue_column: print("[Warning] Glue column not found!")

    #----------------------------------------------------------------------------------------------------
    # main methods
    #----------------------------------------------------------------------------------------------------
    def create_directories(self):
        """ Create folders based on CERN ID """
        print("[INFO] 建立資料夾：")
        for index, row in self.df.iterrows():
            sub_folder = os.path.join(self.base, row['ID'])
            os.makedirs(sub_folder, exist_ok=True)
            print(f"A folder has been created: {sub_folder}")
        print("")

    def create_documents(self):
        """ Create documents """
        print("[INFO] 建立docx文件：")
        for index, row in self.df.iterrows():
            self._create_quality_control_doc(row)

    #----------------------------------------------------------------------------------------------------
    # auxiliary modules
    #----------------------------------------------------------------------------------------------------
    def _create_quality_control_doc(self, row):
        self.cernID = row['ID']
        self.folder = os.path.join(self.base, self.cernID)
        self.gdoc = row['filename+ID'] + '.docx'
        self.output_file = os.path.join(self.folder, self.gdoc)

        self.doc = docx.Document()
        self._set_page_margins()
        self._add_title(row)
        self._add_first_visual_inspection(row)
        self.doc.add_page_break()
        self._add_title(row)
        self._add_second_visual_inspection(row)
        self.doc.save(self.output_file)
        print(f"Document has been saved as {self.output_file}")

    #----------------------------------------------------------------------------------------------------
    # Load-data related
    #----------------------------------------------------------------------------------------------------
    def _check_folder(self): # TODO: return error if folder/csv does not exist
        # 確保目標資料夾存在
        if not os.path.exists(self.base):
            os.makedirs(self.base)
            print(f'新增 {self.base}')

        # List all files in the directory
        files = os.listdir(self.base)
        for f in files:
            full_path = os.path.join(self.base, f)
            file_size = os.path.getsize(full_path)
            # print(f"- {f} ({file_size} bytes)")

        # Check for specific file
        if os.path.exists(self.csv):
            print(f"[INFO] Found target file: {self.filename}")
            print(f"Full path: {self.csv}")
            print(f"File size: {os.path.getsize(self.csv)} bytes")
            print("")
        else:
            print(f"\n[ERROR] Target file does not exist: {self.filename}")

    def _read_and_process_csv(self):
        """
        Read and process the CSV file to extract relevant information
        """
        try:
            # Skip the first two rows and use the third row as headers
            df = pd.read_csv(self.csv, skiprows=2, header=0) # 刪除前兩行
            df = df.dropna(subset=['User'])  # 刪除 User 欄位為空的列
            df = df.fillna('') # replace all NaN with empty strings
            df.columns = [str(col).strip().replace('\n', ' ') for col in df.columns]
            # self._inspect_contents(df)
            return df

        except Exception as e:
            print(f"Error processing file: {str(e)}")
            return None

    def _inspect_contents(self, df):
        # Print the column names to verify
        print("\nColumn names:")
        for i, col in enumerate(df.columns):
            print(f"{i}: '{col}'")

        # Print the first few rows of data
        keywords = ['ID', 'Accept?', 'filename+ID', 'image link']
        print("\nFirst few rows of data (selected columns):")
        print(df[keywords].head())

    def _find_column_by_keyword(self, keyword):
        """
        Find column name that contains the given keyword
        """
        matching_cols = [col for col in self.df.columns if keyword.lower() in col.lower()]
        return matching_cols[0] if matching_cols else None

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

    def _add_empty_spaces(self, paragraph, Nspaces=4, color=None):
        space = ' '
        underlined_spaces = paragraph.add_run(space * Nspaces)
        underlined_spaces.font.color.rgb = color or self.black

    def _add_title(self, row):
        title = self.doc.add_paragraph("Hexaboard 8\"V3 HD-FUll-HB-V2.2 Quality control traveler document V3")
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_format = title.runs[0].font
        title_format.size = Pt(14)
        title_format.bold = True

        info = [
            ("User:", str(row['User']) if pd.notna(row['User']) else ""),
            ("Date:", str(row['Date']) if pd.notna(row['Date']) else ""),
            ("Version:", str(row['Version']) if pd.notna(row['Version']) else ""),
            ("Manufacturer:", str(row['Manufacturer']) if pd.notna(row['Manufacturer']) else ""),
            ("Batch:", str(row['Batch']) if pd.notna(row['Batch']) else ""),
            ("ID:", str(row['ID']) if pd.notna(row['ID']) else "")
        ]

        for i, (key, value) in enumerate(info):
            if i % 3 == 0:
                p = self.doc.add_paragraph()

            run = p.add_run(key)
            run.font.color.rgb = self.black
            self._add_underlined_spaces(p, 2)

            run = p.add_run(value)
            run.font.underline = WD_UNDERLINE.THICK
            run.font.color.rgb = self.blue
            self._add_underlined_spaces(p, 4 - (i // 5))

    def _add_formatted_table(self, row):
        # 添加新的1x2表格
        new_table = self.doc.add_table(rows=1, cols=2)
        new_table.style = 'Normal Table'

        # Add chip ID & chip map
        for i, cell in enumerate(new_table.rows[0].cells):
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i==0:
                # cell.width = Inches(2)
                paragraph = cell.add_paragraph()
                paragraph.add_run(f"{row['p2_Chip ID']}")
            elif i==1:
                # cell.width = Inches(4)
                self._add_image_to_cell(cell, row.get('p2_Chip location map link', ''))

    def _add_image_to_cell(self, cell, image_link, default_width=3):
        """
        Adds an image to a table cell with error handling.

        Args:
            cell: The table cell to add the image to
            image_link: The path to the image
            default_width: Width in inches for the image (default: 3)

        Returns:
            bool: True if image was added successfully, False otherwise
        """
        paragraph = cell.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if not image_link or not image_link.strip():
            run = paragraph.add_run("No image available")
            run.italic = True
            return False

        try:
            image_path = os.path.join(self.base, image_link)
            if not os.path.exists(image_path):
                run = paragraph.add_run("Image file not found")
                run.italic = True
                run.font.color.rgb = RGBColor(255, 0, 0)
                return False

            run = paragraph.add_run()
            run.add_picture(image_path, height=Inches(default_width))
            return True

        except Exception as e:
            # Log the error if needed
            print(f"Error adding image: {str(e)}")
            run = paragraph.add_run("Error loading image")
            run.italic = True
            run.font.color.rgb = RGBColor(255, 0, 0)
            return False

    def _add_customized_paragraph(self, paragraph, item, value, spaces):
        useEmptySpace = (spaces[0]==0) and (spaces[1]==0)

        paragraph.add_run(f"{item} ")
        self._add_underlined_spaces(paragraph, spaces[0])
        run = paragraph.add_run(value)
        if useEmptySpace is False: run.font.underline = WD_UNDERLINE.THICK
        run.font.color.rgb = self.blue
        self._add_underlined_spaces(paragraph, spaces[1])
        if useEmptySpace: self._add_empty_spaces(paragraph)

    def _add_first_visual_inspection(self, row):
        self.doc.add_heading("1st Visual Inspection – Bare PCB", level=1)

        inspection_items = [
            ("General comments:"       , f"{row['General comments']}"         , 2 , [34, 0, 42]),
            ("Flatness:"               , f"{row['Flatness']}"                 , 1 , [2  , 2])  ,
            ("Comments:"               , f"{row['Comments']}"                 , 0 , [25 , 0])  ,
            ("Thickness measurements:" , f"{row['Thickness measurements']}mm" , 1 , [4  , 22]) ,
            ("Plating (BGA):"          , f"{row['Plating (BGA)']}"            , 1 , [4  , 8])  ,
            ("Plating (Holes):"        , f"{row['Plating (Holes)']}"          , 0 , [4  , 8])  ,
            ("Soldermask alignment:"   , f"{row['Soldermask alignment']}"     , 1 , [4  , 26]) ,
            ("Glue problems?"          , f"{row['Glue problems?']}"           , 1 , [7  , 26]) ,
            ("Test coupons (observations, continuity measurements etc.):", f"{row['Test coupons (observations, continuity measurements etc.)']}" , 2, [4, 10, 42]),
            ("Accept?"                 , f"{row['Accept?']}"                  , 1 , [4  , 12]),
        ]

        for item, value, nLines, spaces in inspection_items:
            if 'Accept' in item:
                image_path = os.path.join(self.base, row['image link'])
                paragraph = self.doc.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = paragraph.add_run()
                run.add_picture(image_path, height=Inches(3))
                # print(f'[INFO] add picture: {image_path}')

            if nLines > 0: p = self.doc.add_paragraph()
            self._add_customized_paragraph(p, item, value, spaces)

            if nLines == 2:
                p = self.doc.add_paragraph()
                self._add_underlined_spaces(p, spaces[2])

    def _add_second_visual_inspection(self, row):
        self.doc.add_heading("2nd Visual Inspection – Assembled PCB", level=1)

        # Assembling data
        assembled_items = [
            ("General comments:"    , f"{row['p2_General comments']}"    , 1, [4, 29]),
            ("Flatness:"            , f"{row['p2_Flatness']}"            , 1, [4, 30]),
            ("HGCROC type:"         , f"{row['p2_HGCROC type']}"         , 1, [4,  8]),
            ("HGCROC rotation:"     , f"{row['p2_HGCROC rotation']}"     , 0, [4,  8]),
            ("Connectors:"          , f"{row['p2_Connectors']}"          , 1, [4,  8]),
            ("Resistors/capacitors:", f"{row['p2_Resistors/capacitors']}", 0, [4,  8]),
        ]

        for item, value, nLines, spaces in assembled_items:
            if nLines > 0: p = self.doc.add_paragraph()
            self._add_customized_paragraph(p, item, value, spaces)

        # Table for chip ID and location map
        self.doc.add_paragraph()
        self._add_formatted_table(row)

        # Functional tests
        self.doc.add_heading("Functional Tests", level=1)
        functional_tests = [
            ("Power-on current:" , f"{row['p2_Power-on current']}" , 1, [4, 10]),
            ("Configured OK:"    , f"{row['p2_Configured OK']}"    , 0, [0,  0]),
            ("Operating current:", f"{row['p2_Operating current']}", 0, [4, 10]),
            ("DAQ lines OK: "    , f"{row['p2_DAQ lines OK']}"     , 1, [0,  0]),
        ]

        for item, value, nLines, spaces in functional_tests:
            if nLines > 0: p = self.doc.add_paragraph()
            self._add_customized_paragraph(p, item, value, spaces)

# Usage example
if __name__ == "__main__":
    folder, filename = 'autoDoc', 'test.csv'
    generator = QualityControlDocGenerator(folder, filename)
    generator.create_directories()
    generator.create_documents()
