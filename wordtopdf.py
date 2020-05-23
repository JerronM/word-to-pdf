import os
    import comtypes.client
    import time

    wdFormatPDF = 17

    # path of file
    in_file=r'absolute path of input docx file 1'
    out_file=r'absolute path of output pdf file 1'

    in_file2=r'absolute path of input docx file 2'
    out_file2=r'absolute path of outputpdf file 2'

    in_file3=r'absolute path of input docx file 3'
    out_file3=r'absolute path of output pdf file 3'

    in_file4=r'absolute path of input docx file 4'
    out_file4=r'absolute path of outputpdf file 4'
    
    in_file5=r'absolute path of input docx file 5'
    out_file5=r'absolute path of output pdf file 5'

    in_file6=r'absolute path of input docx file 6'
    out_file6=r'absolute path of outputpdf file 6'
    
    in_file7=r'absolute path of input docx file 7'
    out_file7=r'absolute path of output pdf file 7'

    in_file8=r'absolute path of input docx file 8'
    out_file8=r'absolute path of outputpdf file 8'
    
    in_file9=r'absolute path of input docx file 9'
    out_file9=r'absolute path of output pdf file 9'

    in_file10=r'absolute path of input docx file 10'
    out_file10=r'absolute path of outputpdf file 10'
    
    
    # print out filenames
    print in_file
    print out_file
    print in_file2
    print out_file2
    print in_file3
    print out_file3
    print in_file4
    print out_file4
    print in_file5
    print out_file5
    print in_file6
    print out_file6
    print in_file7
    print out_file7
    print in_file8
    print out_file8
    print in_file9
    print out_file9
    print in_file10
    print out_file10

    # create COM object
    word = comtypes.client.CreateObject('Word.Application')
    # key point 1: make word visible before open a new document
    word.Visible = True
    # key point 2: wait for the COM Server to prepare well.
    time.sleep(3)

    # convert docx file 1 to pdf file 1
    doc=word.Documents.Open(in_file) # open docx file 1
    doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
    doc.Close() # close docx file 1
    word.Visible = False
    # convert docx file 2 to pdf file 2
    doc = word.Documents.Open(in_file2) # open docx file 2
    doc.SaveAs(out_file2, FileFormat=wdFormatPDF) # conversion
    doc.Close() # close docx file 2   
    word.Quit() # close Word Application 
    # convert docx file 3 to pdf file 3
    doc=word.Documents.Open(in_file) # open docx file 3
    doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
    doc.Close() # close docx file 4
    word.Visible = False
    # convert docx file 4 to pdf file 4
    doc=word.Documents.Open(in_file) # open docx file 4
    doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
    doc.Close() # close docx file 4
    word.Visible = False
    # convert docx file 5 to pdf file 5
    doc=word.Documents.Open(in_file) # open docx file 5
    doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
    doc.Close() # close docx file 5
    word.Visible = False
    # convert docx file 6 to pdf file 6
    doc=word.Documents.Open(in_file) # open docx file 6
    doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
    doc.Close() # close docx file 6
    word.Visible = False
    # convert docx file 7 to pdf file 7
    doc=word.Documents.Open(in_file) # open docx file 7
    doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
    doc.Close() # close docx file 7
    word.Visible = False
    # convert docx file 8 to pdf file 8
    doc=word.Documents.Open(in_file) # open docx file 8
    doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
    doc.Close() # close docx file 8
    word.Visible = False
    # convert docx file 9 to pdf file 9
    doc=word.Documents.Open(in_file) # open docx file 9
    doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
    doc.Close() # close docx file 9
    word.Visible = False
    # convert docx file 10 to pdf file 10
    doc=word.Documents.Open(in_file) # open docx file 10
    doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
    doc.Close() # close docx file 10
    word.Visible = False
    
    
