Imports System.ComponentModel
Imports System.IO
Imports System.Xml
Imports FastReport.Export.Dbf

Public Class clsReport
    Private ds As DataSet
    Private xmlFile, xsdFile, rptFile As String

    Sub New()
        xmlFile = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fastreport", "itemlist.xml")
        xsdFile = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fastreport", "itemlist.xsd")
        rptFile = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fastreport", "ProductionReport.frx")
    End Sub
    Public Function ShowReport(ByRef blRecord As BindingList(Of clsRecord)) As Boolean
        Dim isInited As Boolean = InitTheDataSet()
        If isInited = False Then
            Return False
        End If
        Dim isFilled As Boolean = FillTheDataSet(blRecord)
        If isFilled = False Then
            Return False
        End If
        Dim rpt As New FastReport.Report
        rpt.Load(rptFile)
        rpt.RegisterData(ds)
        rpt.Show()
        Return True
    End Function

    Private Function InitTheDataSet() As Boolean
        Try
            Dim strXmlContent As String = System.IO.File.ReadAllText(xmlFile)
            Dim sr As StringReader = New StringReader(strXmlContent)
            Dim xtr As XmlTextReader = New XmlTextReader(sr)
            ds = New DataSet
            'ds.DataSetName = "ds"
            ds.ReadXml(xtr)
            ds.ReadXmlSchema(xsdFile)
            Return True
        Catch ex As Exception
            ' modCommon.WriteErrEvtLog("clsReport.InitTheDataSet: " & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Private Function FillTheDataSet(ByRef lst As BindingList(Of clsRecord)) As Boolean
        If lst Is Nothing Then
            Return False
        End If
        If ds IsNot Nothing Then
            ds.Clear()
        End If
        For Each obj As clsRecord In lst
            Dim _row As DataRow = ds.Tables("item").NewRow()
            _row("objectId") = obj.objectId
            _row("itemnum") = obj.itemnum
            _row("itemname") = obj.itemname
            _row("itemname2") = obj.itemname2
            _row("itemuom") = obj.itemuom
            _row("itemstandardweight") = obj.itemstandardweight
            _row("itemprice") = obj.itemprice
            _row("weightingram") = obj.weightingram
            _row("sellingprice") = obj.sellingprice
            _row("barcode") = obj.barcode
            _row("packingdt") = obj.packingdt
            ds.Tables("item").Rows.Add(_row)
        Next
        Return True
    End Function

End Class
