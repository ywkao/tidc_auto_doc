from tidc_auto_doc import autoDocCreater

tester = autoDocCreater.QualityControlDocGenerator('autoDoc', 'test.csv')

print(tester.create_directories())
