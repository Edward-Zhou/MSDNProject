Imports Newtonsoft.Json
Imports System.Net.Http
Imports System.Text

Public Class Form1
    Dim DSSet As New DataSet

    Private Sub DisplayInDGV()
        'Dim SQLSet As String
        'Dim DASet As New OleDb.OleDbDataAdapter
        'Dim connection As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\v-tazho\Documents\Test1.accdb;Persist Security Info=False;")
        'SQLSet = "Select * From Sheet2"
        'DASet = New OleDb.OleDbDataAdapter(SQLSet, connection)
        'DSSet.Clear()
        'DASet.Fill(DSSet, "DSSetHere")
        'With DGVSetView
        '    .Refresh()
        '    .AutoGenerateColumns = False   'This line must be placed before assigning the datasource to the datagridview'
        '    .DataSource = Nothing
        '    .DataSource = DSSet.Tables(0)
        '    .Update()

        '    .Columns(0).DataPropertyName = DSSet.Tables(0).Columns(0).ToString
        '    .Columns(1).DataPropertyName = DSSet.Tables(0).Columns(1).ToString
        '    .Columns(2).DataPropertyName = DSSet.Tables(0).Columns(2).ToString
        'End With
        'For ItemRow As Integer = 0 To DGVSetView.Rows.Count - 1
        '    DGVSetView.Rows(ItemRow).Cells(3).Value = DGVSetView.Rows(ItemRow).Cells(2).Value
        'Next

        'DGVSetView.Update()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DisplayInDGV()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'Test1DataSet.Sheet2' table. You can move, or remove it, as needed.
        'Me.Sheet2TableAdapter.Fill(Me.Test1DataSet.Sheet2)
        DGVSetView.Rows.Add("T1", "T2", "T3", "T4")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DSSet.Clear()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
    Public StoreResponse As String
    Public Class JSON_postStoreInfo
        Public StoreID As String
        Public POSTransId As String
    End Class
    Public Class JSON_resultStorePos
        'This is the response received from the intial POST for the transaction  - /v2/transaction
        Public code As String
        Public message As String
        Public version As String
        Public transaction_id As String
    End Class
    Private Async Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'declare an instance of JSON_postStoreInfo
        Dim initObj As New JSON_postStoreInfo

        'Assign values for testing
        With initObj
            .StoreID = 45
            .POSTransId = "POS1"
        End With

        'This takes in an object of the JSON_postStoreInfo class 
        'and converts it to JSON to send 

        'Method one:
        Dim JsonString As String = String.Empty
        JsonString = JsonConvert.SerializeObject(initObj)
        'Immediate window:
        ' JsonString
        '{""StoreID"":""45"",""POSTransId"":""POS1""}"
        'Looks like I am getting extra quotes above and the output is a sort of json already 
        'if the extra quotes weren't there 

        'Method two:
        Dim storeHttpContent As StringContent
        storeHttpContent = New StringContent(JsonConvert.SerializeObject(initObj), Encoding.UTF8, "application/json")
        Dim aClient As New HttpClient()
        Dim aResponse As HttpResponseMessage = Await aClient.PostAsync("http://localhost/WCFRESTJson/Service1.svc/PostJson", storeHttpContent)
        If (aResponse.IsSuccessStatusCode) Then
            'this gets the response from remote site
            StoreResponse = aResponse.ToString
            Dim responseContent = aResponse.Content.ReadAsStringAsync().Result

            Dim result As New JSON_resultStorePos
            result = JsonConvert.DeserializeObject(Of JSON_resultStorePos)(responseContent)
            Dim a = result.code
        Else
            'show the response status code 
            Dim failureMsg = "HTTP Status: " + aResponse.StatusCode.ToString() + " – Reason: " + aResponse.ReasonPhrase
            StoreResponse = aResponse.ToString
        End If
        'Get error saying New cannot be used on a class that is declared "MustInherit"

        'So i need to send two values as JSON to a remote site

    End Sub
    Private Async Function SendStoreInfo(Content As HttpContent) As Task(Of HttpResponseMessage)
        Dim theUri As New Uri("http://www.remotesite.net")
        Dim storeHttpContent As HttpContent = Content
        'storeHttpContent = (Content, UnicodeEncoding.UTF8, "application/json")
        Dim aClient As New HttpClient()
        'Dim theContent As New StringContent(SR.ReadToEnd(), System.Text.Encoding.UTF8, "application/json")
        ' Dim theContent As New StringContent(Content, System.Text.Encoding.UTF8, "application/json")
        'Post the data 
        Dim aResponse As HttpResponseMessage = Await aClient.PostAsync(theUri, storeHttpContent)
        If (aResponse.IsSuccessStatusCode) Then
            'this gets the response from remote site
            StoreResponse = aResponse.ToString
        Else
            'show the response status code 
            Dim failureMsg = "HTTP Status: " + aResponse.StatusCode.ToString() + " – Reason: " + aResponse.ReasonPhrase
            StoreResponse = aResponse.ToString
        End If
    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim p As String = "123 /asd"
        Dim arr As Array = Split(p, "/")
        Dim chr As Char
        Dim result As String = ""
        For Each a As String In arr
            a = a.Trim()
            chr = a(a.Length - 1)
            result = result & "/" & chr
        Next
        result = result.TrimStart("/")
        MessageBox.Show(result)
    End Sub

    Private Sub DGVSetView_UserAddedRow(sender As Object, e As DataGridViewRowEventArgs) Handles DGVSetView.UserAddedRow
        'MessageBox.Show("DGVSetView_UserAddedRow")
        Dim rowCount As Integer
        Dim result2 As String
        Dim result3 As String
        rowCount = DGVSetView.Rows.Count
        result2 = DGVSetView.Rows(rowCount - 3).Cells(1).Value
        result3 = DGVSetView.Rows(rowCount - 3).Cells(2).Value
        DGVSetView.Rows(e.Row.Index - 1).Cells(1).Value = result2
        DGVSetView.Rows(e.Row.Index - 1).Cells(2).Value = result3
    End Sub
End Class
