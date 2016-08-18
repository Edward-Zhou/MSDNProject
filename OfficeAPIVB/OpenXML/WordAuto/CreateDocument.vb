Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Public Class CreateDocument
    Public Sub CreateDocFromTem()
        Dim sFile As String = "D:\OfficeDev\OpenXML\Word\Hello World.dotx"
        If File.Exists(sFile.Replace("dotx", "docx")) Then
            File.Delete(sFile.Replace("dotx", "docx"))
        End If
        File.Copy(sFile, sFile.Replace("dotx", "docx"))
        Dim uniEncoding As New UnicodeEncoding()
        Dim fs As New FileStream(sFile, FileMode.Open, FileAccess.Read)
        Dim templateStream As New MemoryStream()
        fs.CopyTo(templateStream)


        ' Modify the document to ensure it is correctly marked as a Document and not Template
        Using document As WordprocessingDocument = WordprocessingDocument.Open(templateStream, True)
            document.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
        End Using

        ' Add the XML into the document and save to the correct location.               
        File.WriteAllBytes(sFile.Replace("dotx", "docx"), templateStream.ToArray())

    End Sub

End Class
