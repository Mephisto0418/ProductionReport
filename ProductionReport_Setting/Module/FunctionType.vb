Module FunctionType

    Public FunctionTypeList As New List(Of String)

    Public Sub FunctionParameterLoad()
        FunctionTypeList.Clear()
        '自動代入欄位設定種類
        FunctionTypeList.Add("OP參數")
        FunctionTypeList.Add("SPC")
        FunctionTypeList.Add("愉進系統")
        FunctionTypeList.Add("欄位間計算")
        FunctionTypeList.Add("IPQC")
        FunctionTypeList.Add("其他")



    End Sub


    Public Sub AddFunctionType(ByRef CboF_Type As ComboBox)
        CboF_Type.Items.Clear()
        For Each type As String In FunctionTypeList
            CboF_Type.Items.Add(type)
        Next
    End Sub


End Module
