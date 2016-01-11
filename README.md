# excelmatrix
a trivial wrapper for openpyxl to import a matrix from excel

convenience funtions for openpyxl:

    from pprint import pprint
    from  excelmatrix import writeMatrix,readMatrix


    testFilename = "export1.xlsx"
    testSheetname = "Blatt1"

    contentMatrix= [[u'1a', u'1b', u'1c', u'1d'], 
                    [u'-', u'Col_1', u'Col_2', u'Col_3'], 
                    [u'Row_1', u'w1', u'w2', u'u3'], 
                    [u'Row_2', u'x1', u'x2', u'x3'], 
                    [u'Row_3', u'y1', u'y2', u'y3']]


    writeMatrix(testFilename,testSheetname,contentMatrix)
    myContent = readMatrix (testFilename,testSheetname,range="A3:F5")
    pprint (myContent)
    myContent = readMatrix (testFilename,testSheetname,start="A3",cellsPerRow=5)
    pprint (myContent)


    
