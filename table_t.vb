Public Class history_t
    Public Property YAKUJYOU_dt As String '๑่๚
    Public Property name As String 'มฟ
    Public Property code As String 'มฟR[h
    Public Property market As String
    Public Property TORIHIKI As String 'ๆ๘
    Public Property KIGEN As String
    Public Property AZUKARI As String
    Public Property KAZEI As String
    Public Property TANKA_sr As String 'MpEปจฬลฬPฟ
    Public Property volume As Integer 'ส
    Public Property cost_sr As String
    Public Property tax_sr As String
    Public Property UKEWATASHIBI As String
    Public Property money_sr As String '๓nเz->Mp:นv | ปจ:บฝฤเz
    'ฑฑฉ็มHl
    Public Property TANKA_ct As Integer = 0 'ปจฬลฬvZPฟ->ยฯPฟ
    Public Property ent_money As Integer = 0
    Public Property enter_m_ct As Integer = 0
    Public Property ext_money As Integer = 0
    Public Property exit_m_ct As Integer = 0
    Public Property SONEKI As Integer = 0
    Public Property exted_vol As Integer = 0
    Public Property enter_dt As String = ""
    Public Property enter_ct As Integer = 0
    Public Property exit_ct As Integer = 0
    Public Function orderBytype() As Integer
        Select Case True
            Case 0 <= TORIHIKI.IndexOf("MpVK")
                Return 0
            Case 0 <= TORIHIKI.IndexOf("MpVK")
                Return 1
            Case 0 <= TORIHIKI.IndexOf("ป๘")
                Return 2
            Case 0 <= TORIHIKI.IndexOf("ปn")
                Return 3
            Case 0 <= TORIHIKI.IndexOf("ฎปจ")
                Return 4
            Case 0 <= TORIHIKI.IndexOf("ช")
                Return 4
            Case 0 <= TORIHIKI.IndexOf("Mpิฯ")
                Return 5
            Case 0 <= TORIHIKI.IndexOf("Mpิฯ")
                Return 6
            Case 0 <= TORIHIKI.IndexOf("ฎปจ")
                Return 7
            Case Else
                Return 100
        End Select
    End Function
    Public Function remainVolume() As Integer
        Return volume - exted_vol
    End Function 'าฟฬc่
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
        Return If(orderBytype() = 5, True, False) 'Mpิฯ่
    End Function
    Public Function IsGenbutsuEntry() As Boolean
        Return If(orderBytype() = 4 OrElse orderBytype() = 2, True, False) 'ปจ,ป๘
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

Public Class output_trade_t
    Public Property ๆพ๚ As String
    Public Property มฟR[h As String
    Public Property มฟผ As String
    Public Property ๆพฟ As String
    Public Property ๆพส As String
    Public Property ๆพเz As String
    Public Property ๆพ่ฟ As String
    Public Property ๆพม๏ล As String
    Public Property ๆ๘Kp As String
    Public Property ฯ๚ As String
    Public Property ฯฟ As String
    Public Property ฯส As String
    Public Property ฯเz As String
    Public Property ฯ่ฟ As String
    Public Property ฯม๏ล As String
    Public Property นvz As String
End Class
Public Class SONEKI_MEISAI_t
    Public Property มฟR[h As String
    Public Property มฟ As String
    Public Property ๗nvๆม As String
    Public Property exit_dt As String
    Public Property ส As String
    Public Property ๆ๘ As String
    Public Property ๓n๚ As String
    Public Property porฯm As String
    Public Property exit_ct As String
    Public Property enter_dt As String
    Public Property ๆพorVKm As String
    Public Property นvเz As String
End Class
Public Class header_t
    Public SYOUHIN_SHITEI As String
    Public KAISHI_DATE As String
    Public SYURYOU_DATE As String
    Public MEISAI_SUU As String
    Public MEISAI_KAISSHI As String
    Public MEISAI_SYURYOU As String
End Class