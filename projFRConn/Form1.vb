Imports System.ComponentModel
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography

Public Class Form1

    Private WithEvents objFetchData As clsFetchData
    Private objReport As clsReport

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        objReport = New clsReport
        objFetchData = New clsFetchData
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim startdt As String = "11/01/2022 00:00:00"
        'Dim enddt As String = "11/30/2022 00:00:00"
        Dim startdt_1 As DateTime = New Date(2022, 11, 1, 0, 0, 0)
        Dim enddt_1 As DateTime = New Date(2022, 12, 1, 0, 0, 0)
        'Dim isparsed As Boolean = DateTime.TryParse(startdt, startdt_1)
        'Dim isparsed2 As Boolean = DateTime.TryParse(enddt, enddt_1)
        objFetchData.getRangeRecords(startdt_1, enddt_1)

    End Sub

    Private Sub DataDownloadedHandler(isSuccess As Boolean) Handles objFetchData.DataDownloaded
        If isSuccess Then
            objReport.ShowReport(objFetchData.blRecord)
        End If
    End Sub




End Class
