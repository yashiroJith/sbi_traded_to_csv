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
#Region "ハンドラ"
    Private Sub textLog_DragEnter(sender As Object, e As DragEventArgs) Handles textLog.DragEnter
        Dim IsDragFile = e.Data.GetDataPresent(DataFormats.FileDrop)
        Dim changeCursor = Sub()
                               e.Effect = DragDropEffects.Copy
                           End Sub
        If IsDragFile Then changeCursor()
    End Sub 'テキストボックスにドラッグオーバー時の処理

#End Region
#Region "メイン"
    Private Sub textLog_DragDrop(sender As Object, e As DragEventArgs) Handles textLog.DragDrop
        Dim fileName As String() = e.Data.GetData(DataFormats.FileDrop, False) 'ファイル群を取得
        If fileName.Length <= 0 Then Exit Sub 'ファイル数が０なら終了
        Dim IsTextBox = TypeOf sender Is TextBox
        If Not IsTextBox Then Exit Sub 'オブジェクトがテキストボックス以外なら終了


        textLog.AppendText(fileName(0) & vbCrLf) '１番目のファイル名を表示
        Dim header = map_header(fileName(0))
        'maplistの読み込み
        Dim list_traded = map_traded(fileName(0))

        Dim query_date = list_traded.OrderBy(Function(t0) t0.date_trade).ThenBy(Function(t1) t1.orderBytype).ToList

        Dim traded As New List(Of List(Of trade_t)) 'トレード済み
        Dim position As New List(Of trade_t) 'トレード中

        For Each t In query_date
            'textLog.AppendText($"{t.index} : {t.date_trade} : {t.code} : {t.name} : {t.volume} : {t.type_trade} : {t.money}" & vbCrLf)

            'tradedからtのコードを探す。なければ新規作成
            Dim posCount = position.Where(Function(d) d.code).Count
            Dim is_entry_onPosition = t.IsEntry AndAlso 0 < posCount
            Select Case True
                Case Not t.IsEntry AndAlso position.Where(Function(d) d.code = t.code).Count = 0
                    Continue For 'exit_noPosition : もし口座にないexitなら次へ
                Case t.IsEntry AndAlso position.Where(Function(d) d.code = t.code).Count = 0 'newEntry : 口座になくエントリーなら新規エントリー処理
                    Dim newTrade = New trade_t
                    newTrade.code = t.code
                    newTrade.volumeCount = t.volume
                    newTrade.histories.Add(t)
                    position.Add(newTrade)
                    Continue For
                Case is_entry_onPosition 'entry_onPosition
                    Dim onPosition = position.Last(Function(a) a.code = t.code)
                    onPosition.volumeCount += t.volume
                Case Not t.IsEntry AndAlso 0 < position.Where(Function(d) d.code = t.code).Count 'exit_onPosition
                    Dim onPosition = position.Last(Function(a) a.code = t.code)
                    onPosition.volumeCount -= t.volume
                    If onPosition.volumeCount = 0 Then
                        traded.Add(position.Where(Function(a) a.code = t.code).ToList) 'トレード中が0ならトレード済み
                        position.RemoveAll(Function(a) a.code = t.code) '手仕舞い済みを消去
                    End If

            End Select
        Next
        'テスト
        Dim test = 0

    End Sub 'ファイルをドラッグ＆ドロップする処理
#End Region
    Private Function map_traded(_fname As String) As List(Of history_t)
        Dim result = read_from_csv(_fname, 10)
        Dim csv = result.Split(vbCrLf)
        Dim lines As New List(Of history_t)
        Dim _buffer As String = ""
        For i As Integer = 0 To csv.Count - 1
            _buffer = csv(i).Replace(vbLf, "")
            Dim _index() = _buffer.Split(",")
            If _buffer = "" Then Continue For
            Dim traded = New history_t With {.date_trade = _index(0),
                                            .name = _index(1),
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
                                            .money = _index(13),
                                            .index = i}
            lines.Add(traded)
        Next
        Return lines
    End Function
    Private Function map_header(_fname As String) As header_t
        Dim _head = read_from_csv(_fname, 5, 5).Split(",") 'ヘッダー行
        Dim _headed = _head.Select(Function(h) h.Replace(vbCrLf, "")) 'フィルタリング
        Return New header_t With {.SYOUHIN_SHITEI = _headed(0),
                                  .KAISHI_DATE = CDate(_headed(1)).ToShortDateString,
                                  .SYURYOU_DATE = CDate(_headed(2)).ToShortDateString,
                                  .MEISAI_SUU = _headed(3),
                                  .MEISAI_KAISSHI = _headed(4),
                                  .MEISAI_SYURYOU = _headed(5)}
    End Function
    Private Function read_from_csv(_fname As String, _startLine As Integer, Optional _endLine As Integer = 0) As String
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
