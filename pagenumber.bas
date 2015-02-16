Attribute VB_Name = "NewMacros"
Sub PageNumberReset()
    Dim pgNo As Long
    Dim n As Long
    Dim pathName As String
    Dim fileNames
    Dim thisFile As String
    Dim aRange As Range

    ' Specify the path to the document files
    pathName = "\skrips muf 2015\"
    ' Create an array holding the document file names, in sequence
    fileNames = Array("`h_ bab i pendahuluan.docx", "`i_ bab ii landasan teori.docx", "`j_ bab iii analisa sistem berjalan.docx", "`k_ bab iv program aplikasi usulan.docx", "`l_ bab v penutup.docx", "`m_ daftarpustaka.docx", "`n_lampiran.docx")

    pgNo = 0
    For n = 0 To UBound(fileNames)
        thisFile = pathName & fileNames(n)
        Application.Documents.Open (thisFile)
        ActiveDocument.Sections(1).Headers(1).PageNumbers.StartingNumber = pgNo + 1
        Set aRange = ActiveDocument.Range
        aRange.Collapse Direction:=wdCollapseEnd
        aRange.Select
        pgNo = Selection.Information(wdActiveEndAdjustedPageNumber)
        Application.Documents(thisFile).Close Savechanges:=wdSaveChanges
    Next n
End Sub
