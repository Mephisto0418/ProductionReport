Public Class Search_Lot
    Public Property Lot As String
    Public Property Layer As String
    Public Property Face As String
    Public Property Sort As Double

    Public Sub New(lot As String)
        Me.Lot = lot
    End Sub

    Public Sub New(lot As String, layer As String, face As String, Optional sort As Double = 0)
        Me.Lot = lot
        Me.Layer = layer
        Me.Face = face
        Me.Sort = sort
    End Sub
End Class
