Public Class traded_t
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

    Private Function IsEntry() As Boolean
        Dim boo As Boolean = False
        If 0 <= type_trade.IndexOf("現物買") Then boo = True '現物エントリー
        If 0 <= type_trade.IndexOf("新規") Then boo = True '信用エントリー
        Return boo
    End Function

End Class
Public Class output_trade_t
    Public Property date_traded As String
    Public Property code As String
    Public Property name As String
    Public Property volume As String
    Public Property SONEKI As String
    Public Property SYUTOKUHI As String '仕掛け値ｘ数量
    Public Property SYOKEIHI As String
    Public Property date_entry As String
End Class
Public Class header_t
    Public SYOUHIN_SHITEI As String
    Public KAISHI_DATE As String
    Public SYURYOU_DATE As String
    Public MEISAI_SUU As String
    Public MEISAI_KAISSHI As String
    Public MEISAI_SYURYOU As String
End Class