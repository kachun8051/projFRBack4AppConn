Imports System.ComponentModel
Imports System.Net
Imports Newtonsoft

Public Class clsFetchData

    Public Event DataDownloaded(isSuccess As Boolean)

    Public blRecord As BindingList(Of clsRecord)

    Public Function getRangeRecords(fromdate As Date, todate As Date) As Boolean
        ' using TLS 1.2
        ' System.Net.ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType)
        Dim myurl As String = "https://parseapi.back4app.com/classes/Production?" &
            "where=" & getParams(fromdate, todate) & "&order=pickedAt"
        Diagnostics.Debug.WriteLine("Url: " & myurl)
        Dim web As New WebClient
        web.Headers.Add(HttpRequestHeader.Accept, "application/json")
        web.Headers.Add(HttpRequestHeader.ContentType, "application/json")
        web.Headers.Add("X-Parse-Application-Id", "5vUD5SzypdFDZfa7Sxjya1yLliHMAJ52ML3sqBf6")
        web.Headers.Add("X-Parse-REST-API-Key", "sgyDDR9YYlvTfkZv1datnUu75nhnnjqejm2yMFNL")
        web.Encoding = System.Text.Encoding.UTF8
        'Add the event handler here
        AddHandler web.DownloadStringCompleted, AddressOf webClient_DownloadStringCompleted2
        Try
            Dim objuri As Uri = New Uri(myurl)
            web.DownloadStringAsync(objuri)
            Return True
        Catch ex As Exception
            Console.WriteLine("clsRecords.getRangeRecords: " & vbCrLf & ex.Message)
            RaiseEvent DataDownloaded(False)
            Return False
        End Try
    End Function

    Private Sub webClient_DownloadStringCompleted2(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        Console.WriteLine(e.Result)
        Dim lstData As List(Of clsRecord) = Nothing
        Dim jsonResultToDict As Dictionary(Of String, List(Of clsRecord)) =
                Json.JsonConvert.DeserializeObject(Of Dictionary(Of String, List(Of clsRecord)))(e.Result)
        jsonResultToDict.TryGetValue("results", lstData)
        ConvertToBindingList(lstData)
        Console.WriteLine("Event is triggered")
        'objReport.ShowReport(blRecord)
        RaiseEvent DataDownloaded(True)
    End Sub

    Private Function getParams(i_from As Date, i_to As Date) As String

        If isTwoDatesEqual(i_from, i_to) Then
            Return getParam(i_from)
        Else
            Dim jobj_3F As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject()
            jobj_3F.Add("__type", "Date")
            jobj_3F.Add("iso", getISODate(i_from))
            Dim jobj_3T As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject()
            jobj_3T.Add("__type", "Date")
            jobj_3T.Add("iso", getISODate(i_to))
            Dim jobj_1 As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject()
            jobj_1.Add("$gte", jobj_3F)
            jobj_1.Add("$lt", jobj_3T)
            Dim jobj_2 As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject()
            jobj_2.Add("packedAt", jobj_1)
            Dim strOutput As String = jobj_2.ToString(Newtonsoft.Json.Formatting.None)
            Diagnostics.Debug.WriteLine(strOutput)
            Return strOutput
        End If

    End Function

    Private Function getParam(i_date As Date) As String

        Dim jobj_1a As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject()
        jobj_1a.Add("__type", "Date")
        jobj_1a.Add("iso", getISODate(i_date))

        Dim jobj_1b As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject()
        jobj_1b.Add("__type", "Date")
        jobj_1b.Add("iso", getISODate(i_date.AddDays(1)))

        Dim jobj_1 As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject()
        jobj_1.Add("$gte", jobj_1a)
        jobj_1.Add("$lt", jobj_1b)
        Dim jobj_2 As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject()
        jobj_2.Add("packedAt", jobj_1)
        Dim strOutput As String = jobj_2.ToString(Json.Formatting.None)
        Diagnostics.Debug.WriteLine(strOutput)
        Return strOutput
    End Function

    Private Function getISODate(i_date As Date) As String
        Dim timeZoneHour As Int32 = getTimeZone()
        Dim strdt As String = i_date.Year.ToString.PadLeft(4, "0"c) & "-" &
            i_date.Month.ToString.PadLeft(2, "0"c) & "-" &
            i_date.Day.ToString.PadLeft(2, "0"c)
        Dim dt As Date = DateTime.Parse(strdt)
        dt = dt.AddHours(-timeZoneHour)
        Dim dt_1 As String = dt.ToString("yyyy-MM-dd HH:mm:ss").Replace(" ", "T") & "Z"
        Diagnostics.Debug.WriteLine("ISO Date: " & dt_1)
        Return dt_1
    End Function

    Private Function isTwoDatesEqual(dt1 As Date, dt2 As Date) As Boolean
        If dt1.Year <> dt2.Year Then
            Return False
        End If
        If dt1.Month <> dt2.Month Then
            Return False
        End If
        If dt1.Day <> dt2.Day Then
            Return False
        End If
        Return True
    End Function

    ' Return positive or negative number
    Private Function getTimeZone() As Int32
        Dim tzi As TimeZoneInfo = TimeZoneInfo.Local
        Dim offset As TimeSpan = tzi.BaseUtcOffset
        Return offset.Hours
    End Function

    Private Sub ConvertToBindingList(ByRef lst As List(Of clsRecord))
        If blRecord Is Nothing Then
            blRecord = New BindingList(Of clsRecord)
        Else
            blRecord.Clear()
        End If
        If lst Is Nothing Then
            Return
        End If
        For Each obj As clsRecord In lst
            blRecord.Add(obj)
        Next

    End Sub

End Class
