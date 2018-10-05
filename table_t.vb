Public Class history_t
    Public Property date_trade As String
    Public Property name As String
    Public Property code As String
    Public Property market As String
    Public Property type_trade As String
    Public Property limit_margin As String
    Public Property AZUKARI As String
    Public Property KAZEI As String
    Public Property volume As String
    Public Property price As String
    Public Property cost As String
    Public Property tax As String
    Public Property UKEWATASHIBI As String
    Public Property money As String
    Public Property index As Integer
    Public Property outdex As Integer
    Public Property volumeCount As Integer
    Public Function orderBytype() As Integer
        Select Case True
            Case type_trade = "株式現物買"
                Return 0
            Case type_trade = "株式現物売"
                Return 1
            Case type_trade = "信用新規買"
                Return 2
            Case type_trade = "現引"
                Return 3
            Case type_trade = "信用新規売"
                Return 4
            Case type_trade = "信用返済売"
                Return 5
            Case type_trade = "信用返済買"
                Return 6
            Case Else
                Return 100
        End Select

    End Function

    Public Function IsEntry() As Boolean
        Dim boo As Boolean = False
        If 0 <= type_trade.IndexOf("現物買") Then boo = True '現物エントリー
        If 0 <= type_trade.IndexOf("新規") Then boo = True '信用エントリー
        Return boo
    End Function
    Public Function IsExit() As Boolean
        Dim boo As Boolean = False
        If 0 <= type_trade.IndexOf("株式現物売") Then boo = True
        If 0 <= type_trade.IndexOf("返済") Then boo = True
        Return boo
    End Function
    Public Function IsGENBIKI() As Boolean
        Return 0 <= type_trade.IndexOf("現引")
    End Function
End Class

Public Class trade_t '　仕掛け -> 手仕舞い -> 1取引
    Public code As String '銘柄コード
    Public volumeCount As Integer '現在保有している株数
    Public histories As New List(Of history_t)

End Class

Public Class output_trade_t
    Public date_traded As String
    Public code As String
    Public name As String
    Public volume As String
    Public type_trade As String
    Public SONEKI As String
    Public SYUTOKUHI As String '仕掛け値ｘ数量
    Public tax As String
    Public cost As String
    Public money As String
    '    Public SYOKEIHI As String
    Public date_entry As String
End Class
Public Class header_t
    Public SYOUHIN_SHITEI As String
    Public KAISHI_DATE As String
    Public SYURYOU_DATE As String
    Public MEISAI_SUU As String
    Public MEISAI_KAISSHI As String
    Public MEISAI_SYURYOU As String
End Class