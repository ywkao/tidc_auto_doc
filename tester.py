import os
from autoDocCreater import QualityControlDocGenerator

folder = 'autoDoc'
csv = '(測試)操作test版本.csv'
# csv = 'autoDoc/V3-HD-Top\ Hexaboard-V1.0_空板\ 檢查清單\ -\ 檢查清單.csv'

csv = [f for f in os.listdir(folder) if f.endswith('.csv') and "V3" in f]
csv = csv[0]
print(csv)

# tester = QualityControlDocGenerator(folder, csv)

tester = QualityControlDocGenerator(target_folder=folder, filename=csv, drive='.', prefix='./')
tester.create_directories()
tester.move_photos()
tester.create_documents()
tester.move_docx()

# tester.move_back_photos()
