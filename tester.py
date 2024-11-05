from autoDocCreater import QualityControlDocGenerator

folder = 'autoDoc'
csv = '(測試)操作test版本.csv'

tester = QualityControlDocGenerator(folder, csv, './')
tester.create_directories()
tester.create_documents()
