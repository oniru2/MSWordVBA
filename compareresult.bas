Sub compare()
'
' compare Macro
'
'

Dim oDoc1 As Document
Dim oDoc2 As Document

Set oDoc1 = Documents.Open("C:\Users\Oniru\Documents\Compare\testrtf1.rtf")
Set oDoc2 = Documents.Open("C:\Users\Oniru\Documents\Compare\testrtf2.rtf")

Application.CompareDocuments oDoc1, oDoc2, wdCompareDestinationNew, , , , , , True, True

oDoc1.Close
oDoc2.Close
ActiveDocument.SaveAs2 "C:\Users\Oniru\Documents\Compare\result.rtf", wdFormatRTF
ActiveDocument.Close

End Sub
