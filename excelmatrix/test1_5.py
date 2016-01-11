from pprint import pprint
from  excelmatrix import writeMatrix,readMatrix


testFilename = "export1.xlsx"
testSheetname = "Blatt1"
contentMatrix= [[u'1a', u'1b', u'1c', u'1d'], 
                [u'2a', u'2b', u'2c', u'2d'], 
                [u'3a', u'3b', u'3c', u'3d'], 
                [u'4a', u'4b', u'4c', u'4d'], 
                [u'5a', u'5b', u'5c', u'5d']]


writeMatrix(testFilename,testSheetname,contentMatrix)
myContent = readMatrix (testFilename,testSheetname,range="A1:D5")
pprint (myContent)
myContent = readMatrix (testFilename,testSheetname,start="A3",cellsPerRow=4)
pprint (myContent)
