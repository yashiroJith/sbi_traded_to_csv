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
            Case type_trade = "����������"
                Return 0
            Case type_trade = "����������"
                Return 1
            Case type_trade = "�M�p�V�K��"
                Return 2
            Case type_trade = "����"
                Return 3
            Case type_trade = "�M�p�V�K��"
                Return 4
            Case type_trade = "�M�p�ԍϔ�"
                Return 5
            Case type_trade = "�M�p�ԍϔ�"
                Return 6
            Case Else
                Return 100
        End Select

    End Function

    Public Function IsEntry() As Boolean
        Dim boo As Boolean = False
        If 0 <= type_trade.IndexOf("������") Then boo = True '�����G���g���[
        If 0 <= type_trade.IndexOf("�V�K") Then boo = True '�M�p�G���g���[
        Return boo
    End Function
    Public Function IsExit() As Boolean
        Dim boo As Boolean = False
        If 0 <= type_trade.IndexOf("����������") Then boo = True
        If 0 <= type_trade.IndexOf("�ԍ�") Then boo = True
        Return boo
    End Function
    Public Function IsGENBIKI() As Boolean
        Return 0 <= type_trade.IndexOf("����")
    End Function
End Class

Public Class trade_t '�@�d�|�� -> ��d���� -> 1���
    Public code As String '�����R�[�h
    Public volumeCount As Integer '���ݕۗL���Ă��銔��
    Public histories As New List(Of history_t)

End Class

Public Class output_trade_t
    Public date_traded As String
    Public code As String
    Public name As String
    Public volume As String
    Public type_trade As String
    Public SONEKI As String
    Public SYUTOKUHI As String '�d�|���l������
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