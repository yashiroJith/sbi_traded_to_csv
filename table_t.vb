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
            Case type_trade = "�M�p�V�K��"
                Return 0
            Case type_trade = "�M�p�V�K��"
                Return 1
            Case type_trade = "����"
                Return 2
            Case type_trade = "���n"
                Return 3
            Case type_trade = "����������"
                Return 4
            Case type_trade = "�M�p�ԍϔ�"
                Return 5
            Case type_trade = "�M�p�ԍϔ�"
                Return 6
            Case type_trade = "����������"
                Return 7
            Case Else
                Return 100
        End Select
    End Function
    Public Function Is_entry() As Boolean
        Return If(orderBytype() <= 4, True, False)
    End Function
    Public Function IsMeaginSellEntry() As Boolean
        Return If(orderBytype() = 1, True, False) '�M�p���G���g���[
    End Function
    Public Function IsMeaginBuyEntry() As Boolean
        Return If(orderBytype() = 0, True, False) '�M�p���G���g���[
    End Function
    Public Function IsGenbutsuEntry() As Boolean
        Return If(orderBytype() = 4 OrElse orderBytype() = 2, True, False) '������,����
    End Function
    Public Function c_money() As Integer
        Return price * volume
    End Function
    Public Function c_cost() As Integer
        Return CInt(cost) + CInt(tax)
    End Function
End Class

Public Class output_trade_t
    Public Property �擾�� As String
    Public Property �����R�[�h As String
    Public Property ������ As String
    Public Property �擾���� As String
    Public Property �擾���� As String
    Public Property �擾���z As String
    Public Property �擾�萔�� As String
    Public Property �擾����� As String
    Public Property ����K�p As String
    Public Property ���ϓ� As String
    Public Property ���ϊ��� As String
    Public Property ���ϐ��� As String
    Public Property ���ϋ��z As String
    Public Property ���ώ萔�� As String
    Public Property ���Ϗ���� As String
    Public Property ���v�z As String
End Class
Public Class header_t
    Public SYOUHIN_SHITEI As String
    Public KAISHI_DATE As String
    Public SYURYOU_DATE As String
    Public MEISAI_SUU As String
    Public MEISAI_KAISSHI As String
    Public MEISAI_SYURYOU As String
End Class