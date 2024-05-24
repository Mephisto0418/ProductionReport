Public Class Search_Lot
    Public Property Lot As String
    Public Property Layer As String
    Public Property Face As String

    Public Sub New(lot As String)
        Me.Lot = lot
    End Sub

    Public Sub New(lot As String, layer As String, face As String)
        Me.Lot = lot
        Me.Layer = layer
        Me.Face = face
    End Sub
End Class
