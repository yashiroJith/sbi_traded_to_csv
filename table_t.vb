Public Class history_t
    Public Property dt_YAKUJYOU As String '����
    Public Property name As String '����
    Public Property code As String '�����R�[�h
    Public Property market As String
    Public Property TORIHIKI As String '���
    Public Property KIGEN As String
    Public Property AZUKARI As String
    Public Property KAZEI As String
    Public Property TANKA_sr As String '�M�p�E�����̍ŏ��̒P��
    Public Property volume As Integer '����
    Public Property cost_sr As String
    Public Property tax_sr As String
    Public Property UKEWATASHIBI As String
    Public Property money_sr As String '��n���z->�M�p:���v | ����:��č����z
    '����������H�l
    Public Property TANKA_ct As Integer = 0 '�����̍ŏ��̌v�Z�P��->�ϒP��
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
            Case 0 <= TORIHIKI.IndexOf("�M�p�V�K��")
                Return 0
            Case 0 <= TORIHIKI.IndexOf("�M�p�V�K��")
                Return 1
            Case 0 <= TORIHIKI.IndexOf("����")
                Return 2
            Case 0 <= TORIHIKI.IndexOf("���n")
                Return 3
            Case 0 <= TORIHIKI.IndexOf("����������")
                Return 4
            Case 0 <= TORIHIKI.IndexOf("������")
                Return 4
            Case 0 <= TORIHIKI.IndexOf("�M�p�ԍϔ�")
                Return 5
            Case 0 <= TORIHIKI.IndexOf("�M�p�ԍϔ�")
                Return 6
            Case 0 <= TORIHIKI.IndexOf("����������")
                Return 7
            Case Else
                Return 100
        End Select
    End Function
    Public Function remainVolume() As Integer
        Return volume - exted_vol
    End Function '�����҂��̎c�芔��
    Public Function Is_entry() As Boolean
        Return If(orderBytype() <= 4, True, False)
    End Function
    Public Function IsMeaginSellEntry() As Boolean
        Return If(orderBytype() = 1, True, False) '�M�p�V�K��
    End Function
    Public Function IsMeaginBuyEntry() As Boolean
        Return If(orderBytype() = 0, True, False) '�M�p�V�K��
    End Function
    Public Function IsMeaginSellExit() As Boolean
        Return If(orderBytype() = 5, True, False) '�M�p�ԍϔ���
    End Function
    Public Function IsGenbutsuEntry() As Boolean
        Return If(orderBytype() = 4 OrElse orderBytype() = 2, True, False) '������,����
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
    Public Property ���� As String
    Public Property ���� As String
    Public Property �����R�[�h As String
    Public Property �s�� As String
    Public Property ��� As String
    Public Property ���� As String
    Public Property �a�� As String
    Public Property �ې� As String
    Public Property ��萔�� As String
    Public Property ���P�� As String
    Public Property �萔�� As String
    Public Property �Ŋz As String
    Public Property ��n�� As String
    Public Property ��n���z As String
    Public Property �����ϐ��� As String
End Class
Public Class SONEKI_MEISAI_t
    Public Property �����R�[�h As String
    Public Property ���� As String
    Public Property ���n�v��� As String
    Public Property ���� As String
    Public Property ���� As String
    Public Property ��� As String
    Public Property ��n�� As String
    Public Property ���pl���ϋ��z As String
    Public Property ��p As String
    Public Property �擾l�V�K�N���� As String
    Public Property �擾l�V�K���z As String
    Public Property ���v���z As String
End Class
Public Class header_t
    Public SYOUHIN_SHITEI As String
    Public KAISHI_DATE As String
    Public SYURYOU_DATE As String
    Public MEISAI_SUU As String
    Public MEISAI_KAISSHI As String
    Public MEISAI_SYURYOU As String
End Class'SBI���ׂ̌������p