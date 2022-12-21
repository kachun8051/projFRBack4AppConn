Public Class Form1

    Private WithEvents objFetchData As clsFetchData

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = "Generate Report"

        objFetchData = New clsFetchData
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dtFrom As DateTime = DateTimePicker1.Value
        Dim dtTo As DateTime = DateTimePicker2.Value
        Dim dt_1 As DateTime = New Date(2022, 11, 1, 0, 0, 0)
        Dim dt_2 As DateTime = New Date(2022, 11, 30, 0, 0, 0)
        objFetchData.getRangeRecords(dtFrom, dtTo)
    End Sub

    Private Sub DataDownloadedHandler(isSuccess As Boolean) Handles objFetchData.DataDownloaded
        If isSuccess Then
            Dim objReport As clsReport = New clsReport
            With objFetchData
                objReport.setParams(.startDate, .endDate, .rangeType)
            End With
            objReport.ShowReport(objFetchData.blRecord)
        End If
    End Sub



End Class
