Imports System.Text
Imports System.IO
Imports System.Environment

Public Class Form1

    Friend WithEvents textLog As TextBox

#Region "フォーム初期設定"
    Public Sub New()
        init()
    End Sub
    Private Sub init()
        Dim init_form = Sub()
                            Me.Size = New Size(400, 300)
                            Text = "SBI約定履歴パーサー"
                            Me.FormBorderStyle = FormBorderStyle.Sizable
                            Me.MinimumSize = New Size(400, 300)
                        End Sub
        Dim init_textBox = Sub()
                               textLog = New TextBox
                               textLog.AllowDrop = True
                               textLog.Anchor = AnchorStyles.Top OrElse AnchorStyles.Bottom OrElse AnchorStyles.Left OrElse AnchorStyles.Right
                               textLog.Font = New Font("MS UI Gothic", 9.75)
                               textLog.Location = New Point(12, 12)
                               textLog.Multiline = True
                               textLog.ReadOnly = True
                               textLog.ScrollBars = ScrollBars.Both
                               textLog.Size = New Size(360, 237)
                               textLog.Text = "ここに約定履歴CSVファイルをドラッグ＆ドロップしてください"
                               Me.Controls.Add(textLog)
                               textLog.AppendText(vbCrLf & vbCrLf) 'テキストボックスの改行
                           End Sub
        init_form()
        init_textBox()
    End Sub
#End Region
    Private Sub textLog_DragEnter(sender As Object, e As DragEventArgs) Handles textLog.DragEnter
        Dim IsDragFile = e.Data.GetDataPresent(DataFormats.FileDrop)
        Dim changeCursor = Sub()
                               e.Effect = DragDropEffects.Copy
                           End Sub
        If IsDragFile Then changeCursor()
    End Sub 'テキストボックスにドラッグオーバー時の処理

    Private Sub textLog_DragDrop(sender As Object, e As DragEventArgs) Handles textLog.DragDrop
        Dim fileName As String() = e.Data.GetData(DataFormats.FileDrop, False) 'ファイル群を取得
        If fileName.Length <= 0 Then Exit Sub 'ファイル数が０なら終了
        Dim IsTextBox = TypeOf sender Is TextBox
        If Not IsTextBox Then Exit Sub 'オブジェクトがテキストボックス以外なら終了


        textLog.AppendText(fileName(0) & vbCrLf) '１番目のファイル名を表示
        Dim header = map_header(fileName(0))
        'maplistの読み込み
        Dim list_traded = map_traded(fileName(0))

        For Each t In list_traded
            textLog.AppendText($"{t.date_trade} : {t.code} : {t.name}" & vbCrLf)
        Next
    End Sub 'ファイルをドラッグ＆ドロップする処理
    Private Function map_traded(_fname As String) As List(Of traded_t)
        Dim result = read_csv_to_table(_fname, 10)
        Dim csv = result.Split(vbCrLf)
        Dim lines As New List(Of traded_t)
        Dim _buffer As String = ""
        For Each c In csv
            _buffer = c.Replace(vbLf, "")
            Dim _index() = _buffer.Split(",")
            If _buffer = "" Then Continue For
            Dim traded = New traded_t With {.date_trade = _index(0),
                                            .Name = _index(1),
                                            .code = _index(2),
                                            .market = _index(3),
                                            .type_trade = _index(4),
                                            .limit_margin = _index(5),
                                            .AZUKARI = _index(6),
                                            .KAZEI = _index(7),
                                            .volume = _index(8),
                                            .price = _index(9),
                                            .cost = _index(10),
                                            .tax = _index(11),
                                            .UKEWATASHIBI = _index(12),
                                            .money = _index(13)}
            lines.Add(traded)
        Next
        Return lines
    End Function
    Private Function map_header(_fname As String) As header_t
        Dim _head = read_csv_to_table(_fname, 5, 5).Split(",") 'ヘッダー行
        Dim _headed = _head.Select(Function(h) h.Replace(vbCrLf, "")) 'フィルタリング
        Return New header_t With {.SYOUHIN_SHITEI = _headed(0),
                                  .KAISHI_DATE = CDate(_headed(1)).ToShortDateString,
                                  .SYURYOU_DATE = CDate(_headed(2)).ToShortDateString,
                                  .MEISAI_SUU = _headed(3),
                                  .MEISAI_KAISSHI = _headed(4),
                                  .MEISAI_SYURYOU = _headed(5)}
    End Function
    Private Function read_csv_to_table(_fname As String, _startLine As Integer, Optional _endLine As Integer = 0) As String
        Using sr = New StreamReader(_fname, Encoding.GetEncoding("shift_jis"))
            Dim _result As String = ""
            Dim _buffer As String = ""
            Dim i As Integer = 0
            While 0 <= sr.Peek()
                i += 1
                _buffer = sr.ReadLine().Replace("""", "")
                If (_startLine <= i AndAlso i <= _endLine) Or (_startLine <= i AndAlso _endLine = 0) Then _result += _buffer + NewLine
            End While
            Return _result
        End Using
    End Function


End Class
