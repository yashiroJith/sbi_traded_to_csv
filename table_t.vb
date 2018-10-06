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
    Public Property volumeCount As Integer
    Public Property money_enter As String
    Public Property money_exit As String
    Public Property money_SONEKI As String
    Public Property exited_volume As Integer
    Public Function orderBytype() As Integer
        Select Case True
            Case type_trade = "信用新規買"
                Return 0
            Case type_trade = "信用新規売"
                Return 1
            Case type_trade = "現引"
                Return 2
            Case type_trade = "現渡"
                Return 3
            Case type_trade = "株式現物買"
                Return 4
            Case type_trade = "信用返済売"
                Return 5
            Case type_trade = "信用返済買"
                Return 6
            Case type_trade = "株式現物売"
                Return 7
            Case Else
                Return 100
        End Select
    End Function
    Public Function Is_entry() As Boolean
        Return If(orderBytype() <= 4, True, False)
    End Function
    Public Function IsMeaginSellEntry() As Boolean
        Return If(orderBytype() = 1, True, False) '信用売エントリー
    End Function
    Public Function IsMeaginBuyEntry() As Boolean
        Return If(orderBytype() = 0, True, False) '信用買エントリー
    End Function
    Public Function IsGenbutsuEntry() As Boolean
        Return If(orderBytype() = 4 OrElse orderBytype() = 2, True, False) '現物買,現引
    End Function
    Public Function c_money() As Integer
        Return price * volume
    End Function
    Public Function c_cost() As Integer
        Return CInt(cost) + CInt(tax)
    End Function
End Class

Public Class output_trade_t
    Public Property 取得日 As String
    Public Property 銘柄コード As String
    Public Property 銘柄名 As String
    Public Property 取得株価 As String
    Public Property 取得数量 As String
    Public Property 取得金額 As String
    Public Property 取得手数料 As String
    Public Property 取得消費税 As String
    Public Property 取引適用 As String
    Public Property 決済日 As String
    Public Property 決済株価 As String
    Public Property 決済数量 As String
    Public Property 決済金額 As String
    Public Property 決済手数料 As String
    Public Property 決済消費税 As String
    Public Property 損益額 As String
End Class
Public Class header_t
    Public SYOUHIN_SHITEI As String
    Public KAISHI_DATE As String
    Public SYURYOU_DATE As String
    Public MEISAI_SUU As String
    Public MEISAI_KAISSHI As String
    Public MEISAI_SYURYOU As String
End Class