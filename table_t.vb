Public Class history_t
    Public Property dt_YAKUJYOU As String 'ñèú
    Public Property name As String 'Á¿
    Public Property code As String 'Á¿R[h
    Public Property market As String
    Public Property TORIHIKI As String 'æø
    Public Property KIGEN As String
    Public Property AZUKARI As String
    Public Property KAZEI As String
    Public Property TANKA_sr As String 'MpE»¨ÌÅÌP¿
    Public Property volume As Integer 'Ê
    Public Property cost_sr As String
    Public Property tax_sr As String
    Public Property UKEWATASHIBI As String
    Public Property money_sr As String 'ónàz->Mp:¹v | »¨:º½Äàz
    '±±©çÁHl
    Public Property TANKA_ct As Integer = 0 '»¨ÌÅÌvZP¿->ÂÏP¿
    Public Property money_enter As Integer = 0
    Public Property money_enter_ct As Integer = 0
    Public Property money_exit As Integer = 0
    Public Property money_exit_ct As Integer = 0
    Public Property SONEKI As Integer = 0
    Public Property exted_vol As Integer = 0
    Public Property dt_enter As String = ""
    Public Property cost_enter As Integer = 0
    Public Property cost_exit As Integer = 0
    Public Function orderBytype() As Integer
        Select Case True
            Case 0 <= TORIHIKI.IndexOf("MpVK")
                Return 0
            Case 0 <= TORIHIKI.IndexOf("MpVK")
                Return 1
            Case 0 <= TORIHIKI.IndexOf("»ø")
                Return 2
            Case 0 <= TORIHIKI.IndexOf("»n")
                Return 3
            Case 0 <= TORIHIKI.IndexOf("®»¨")
                Return 4
            Case 0 <= TORIHIKI.IndexOf("ª")
                Return 4
            Case 0 <= TORIHIKI.IndexOf("MpÔÏ")
                Return 5
            Case 0 <= TORIHIKI.IndexOf("MpÔÏ")
                Return 6
            Case 0 <= TORIHIKI.IndexOf("®»¨")
                Return 7
            Case Else
                Return 100
        End Select
    End Function
    Public Function remainVolume() As Integer
        Return volume - exted_vol
    End Function 'Ò¿Ìcè
    Public Function Is_entry() As Boolean
        Return If(orderBytype() <= 4, True, False)
    End Function
    Public Function IsMeaginSellEntry() As Boolean
        Return If(orderBytype() = 1, True, False) 'MpVK
    End Function
    Public Function IsMeaginBuyEntry() As Boolean
        Return If(orderBytype() = 0, True, False) 'MpVK
    End Function
    Public Function IsMeaginSellExit() As Boolean
        Return If(orderBytype() = 5, True, False) 'MpÔÏè
    End Function
    Public Function IsGenbutsuEntry() As Boolean
        Return If(orderBytype() = 4 OrElse orderBytype() = 2, True, False) '»¨,»ø
    End Function
    Public Function c_money_shinyo() As Integer
        Return TANKA_sr * volume
    End Function
    Public Function c_cost() As Integer
        If cost_sr = "" OrElse cost_sr = "--" Then cost_sr = 0
        If tax_sr = "" OrElse tax_sr = "--" Then tax_sr = 0
        Return CInt(cost_sr) + CInt(tax_sr)
    End Function
End Class

Public Class patch_t
    Public Property ñèú As String
    Public Property Á¿ As String
    Public Property Á¿R[h As String
    Public Property sê As String
    Public Property æø As String
    Public Property úÀ As String
    Public Property aè As String
    Public Property ÛÅ As String
    Public Property ñèÊ As String
    Public Property ñèP¿ As String
    Public Property è¿ As String
    Public Property Åz As String
    Public Property ónú As String
    Public Property ónàz As String
    Public Property ÏÊ As String
End Class
Public Class SONEKI_MEISAI_t
    Public Property Á¿R[h As String
    Public Property Á¿ As String
    Public Property ÷nvæÁ As String
    Public Property ñèú As String
    Public Property Ê As String
    Public Property æø As String
    Public Property ónú As String
    Public Property plÏàz As String
    Public Property ïp As String
    Public Property æ¾lVKNú As String
    Public Property æ¾lVKàz As String
    Public Property ¹vàz As String
End Class
Public Class header_t
    Public SYOUHIN_SHITEI As String
    Public KAISHI_DATE As String
    Public SYURYOU_DATE As String
    Public MEISAI_SUU As String
    Public MEISAI_KAISSHI As String
    Public MEISAI_SYURYOU As String
End Class'SBI¾×Ì©ª¯p