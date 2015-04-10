Attribute VB_Name = "NewMacros"
Sub PageNumberReset()
    Dim pgNo As Long
    Dim n As Long
    Dim pathName As String
    Dim fileNames
    Dim thisFile As String
    Dim aRange As Range

    ' Specify the path to the document files
    pathName = "..\proposal"
    ' Create an array holding the document file names, in sequence
    fileNames = Array(
	"`g_ bab 1 pendahuluan.docx", 
	"`h_ bab 2 tinjauan pustaka.docx", 
	"`i_ bab 3 metode penelitian.docx", 
	"`j_ jadwal pelaksanaan penelitian.docx", 
	"`k_ daftar pustaka.docx", 
	"`l_ lampiran.docx"
	)

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
