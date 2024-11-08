from autoDocCreater import QualityControlDocGenerator

folder = 'autoDoc'
csv = '(測試)操作test版本.csv'

# tester = QualityControlDocGenerator(folder, csv)

tester = QualityControlDocGenerator(target_folder=folder, filename=csv, drive='.', prefix='./')
tester.create_directories()
tester.move_photos()
tester.create_documents()

tester.move_back_photos()
