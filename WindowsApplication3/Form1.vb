Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    'OLEDB for reading .xls/.xlsx
    Private Const OleSetJet As String = "Microsoft.Jet.OLEDB.4.0"
    Private Const OleSetAce As String = "Microsoft.ACE.OLEDB.12.0"
    Private Const OleVer80 As String = "Excel 8.0"
    Private Const OleVer120 As String = "Excel 12.0 Xml"
    Private getHeader As String
    Private oleSet As String
    Private oleVer As String
    Dim Bobert As String
    Public PricingSheet As DataSet


    Sub Test()

        Dim Filename As String
        Dim xlsx As New Excel.Application
        Dim xlsxWorkbooks As Excel.Workbooks = xlsx.Workbooks
        Dim thisWeekReport As Excel.Workbooks = xlsxWorkbooks.Open(Filename)


    End Sub
    Public Function ReadXls() As DataSet
        Dim returnData As DataSet = New DataSet
        Dim result As DialogResult = OpenFileDialog1.ShowDialog

        If result <> DialogResult.OK Then
            Return Nothing
        End If

        Label1.Text = OpenFileDialog1.FileName
        Dim filePath = OpenFileDialog1.FileName

        Try
            Dim xlsx As New Excel.Application
            Dim xlsxWorkbooks As Excel.Workbooks = xlsx.Workbooks
            Dim thisWeekReport As Excel.Workbooks = xlsxWorkbooks.Open(filePath)

            'For each sheet in the workbook
            'Get the used range 
            'create blank data table, name it the sheet name
            'For each column in the used range
            '   add a column to the data table
            'For each row in the used range
            '   create a new data row
            '   for each column in the used range
            '       new row[column - 1] = cell [row, column]
            '   Add new row to table
            'If the new table has more then 0 rows
            '   add the table to the data set
            'If the data set has more then 0 tables
            '   return the data set, else return nothing



            'MSDN Documentation?


            Return returnData


        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try
    End Function 'ReadXls
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Test()
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        PricingSheet = ReadXls()
    End Sub
End Class
