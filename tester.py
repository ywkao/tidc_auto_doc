from autoDocCreater import QualityControlDocGenerator

folder = 'autoDoc'
csv = 'test.csv'

tester = QualityControlDocGenerator(folder, csv, './')

tester.create_directories()
tester.create_documents()
