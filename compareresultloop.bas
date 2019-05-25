Sub compare()
'
' compare Macro
'
'
Dim doc1 As Document, doc2 As Document

oDoc1 = Dir("C:\Users\Oniru\Documents\Compare\1\")
oDoc2 = Dir("C:\Users\Oniru\Documents\Compare\2\")
While oDoc1 <> "" And oDoc2 <> ""
ChDir "C:\Users\Oniru\Documents\Compare\1\"
Set doc1 = Documents.Open(oDoc1)
ChDir "C:\Users\Oniru\Documents\Compare\2\"
Set doc2 = Documents.Open(oDoc2)
Application.CompareDocuments doc1, doc2, wdCompareDestinationNew, , , , , , True, True
doc1.Close
doc2.Close
ChDir "C:\Users\Oniru\Documents\Compare\result"
ActiveDocument.SaveAs2 FileName:="1.rtf", Fileformat:=wdFormatRTF
ActiveDocument.Close
Wend
End Sub
