Public Class clsRecord
    Public Property objectId As String
    Public Property itemnum As String
    Public Property itemname As String
    Public Property itemname2 As String
    Public Property itemuom As String
    Public Property itemstandardweight As Int32
    Public Property itemprice As Decimal
    Public Property weightingram As Decimal
    Public Property sellingprice As Decimal
    Public Property barcode As String
    Public Property packingdt As String

    'Public Property packingdt As String
    '    Get
    '        Dim tmp As String = packingdt.Substring(0, 19)
    '        Return
    '    End Get
    '    Set(value As String)

    '    End Set
    'End Property

    ' Public Property objProduct As clsProduct

    Public Function mySerialize() As String

    End Function

    Public Function myDeserialize(ByVal json As String) As Boolean

    End Function

End Class