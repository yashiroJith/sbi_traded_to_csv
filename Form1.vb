Imports System.Text
Imports System.IO
Imports System.Environment
Imports CsvHelper

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
        Dim query_date = map_traded(fileName(0)) '日付順に出力
        Dim query_code = orderBy(query_date) '銘柄別に出力

        Dim soneki = Calc_SONEKI(query_code)
        Dim form_soneki = swap_soneki(soneki)
        writeCsv("C:\Users\username0777\Downloads\testData.csv", soneki)
        writeCsv("C:\Users\username0777\Downloads\outData.csv", form_soneki)

    End Sub 'ファイルをドラッグ＆ドロップする処理

    Private Function swap_soneki(soneki As List(Of history_t)) As List(Of output_trade_t)
        Dim output_trade = New List(Of output_trade_t)
        For Each s In soneki
            Dim trade = New output_trade_t
            If s.Is_entry Then
                trade.取得日 = s.date_trade
                trade.銘柄コード = s.code
                trade.銘柄名 = s.name
                trade.取得株価 = s.price
                trade.取得数量 = s.volume
                trade.取得金額 = s.money_enter * -1
                trade.取得手数料 = s.cost
                trade.取得消費税 = s.tax
                trade.取引適用 = s.type_trade
            Else
                trade.銘柄コード = s.code
                trade.銘柄名 = s.name
                trade.取引適用 = s.type_trade
                trade.決済日 = s.date_trade
                trade.決済株価 = s.price
                trade.決済数量 = s.volume
                trade.決済金額 = s.money_exit
                trade.決済手数料 = s.cost
                trade.決済消費税 = s.tax
                trade.損益額 = s.money_SONEKI
            End If
            output_trade.Add(trade)
        Next
        Return output_trade
    End Function

    Private Function Calc_SONEKI(_source As List(Of history_t)) As List(Of history_t)
        Dim retList = New List(Of history_t)
        For Each t In _source
            Dim ret = New history_t
            ret = copyHistory(t)
            Select Case True
                Case 0 <= t.type_trade.IndexOf("信用新規買")
                    ret.tax = ""
                    ret.cost = ""
                    ret.money_enter = ret.c_money * -1
                    ret.money_exit = ""
                    ret.money_SONEKI = ""
                Case 0 <= t.type_trade.IndexOf("信用新規売")
                    ret.tax = ""
                    ret.cost = ""
                    ret.money_enter = ret.c_money
                    ret.money_exit = ""
                    ret.money_SONEKI = ""
                Case 0 <= t.type_trade.IndexOf("現引") '信用分を取り消す
                    ret.money_enter = ret.c_money * -1
                    ret.money_exit = ""
                    ret.money_SONEKI = ""
                    Dim entres = SHINYO_buy_entry(t, retList).Where(Function(s) s.price = ret.price).ToList
                    For Each e In entres '処理済みリスト
                        Dim IsExited = e.volume - e.exited_volume < t.volume
                        If IsExited Then Continue For '残出来高が小さければ次へ
                        e.money_enter = e.price * (e.volume - t.volume) '処理済み金額を差し引く
                        e.exited_volume += t.volume
                    Next
                Case 0 <= t.type_trade.IndexOf("現渡") '信用分を取り消す
                    ret.money_enter = ret.c_money
                    ret.money_exit = ""
                    ret.money_SONEKI = ""
                    Dim entres = SHINYO_sell_entry(t, retList).Where(Function(s) s.price = ret.price).ToList
                    For Each e In entres '処理済みリスト
                        Dim IsExited = e.volume - e.exited_volume < t.volume
                        If IsExited Then Continue For '残出来高が小さければ次へ
                        e.money_enter = e.price * (e.volume - t.volume) '処理済み金額を差し引く
                        e.exited_volume += t.volume
                    Next
                Case 0 <= t.type_trade.IndexOf("株式現物買")
                    ret.money_enter = ret.c_money * -1
                    ret.money_exit = ""
                    ret.money_SONEKI = ""
                Case 0 <= t.type_trade.IndexOf("株式現物売") '以前の現引きか現物買いから損益を計算
                    ret.money_enter = ""
                    Dim entres = GENBUTSU_entry(t, retList)
                    Dim firstTime As Boolean = True
                    For Each e In entres '仕掛けリスト
#Region "手仕舞い条件"
                        Dim has_exit_volume = t.exited_volume < t.volume '手仕舞い残りがある
                        Dim has_entry_volume = 0 < e.volume - e.exited_volume '仕掛け残りがある
                        Dim exit_moreThen_entry = e.volume - e.exited_volume < t.volume '手仕舞いが仕掛け残りより多い

                        Dim entry_moreThen_exit = t.volume <= e.volume - e.exited_volume '仕掛け残り枚数が手仕舞い以上
#End Region
                        If exit_moreThen_entry AndAlso has_entry_volume AndAlso has_exit_volume Then '手仕舞いが仕掛けより多い
                            If firstTime Then
                                ret.money_SONEKI += t.price * e.volume - t.c_cost - e.c_money - e.c_cost '最初だけ手仕舞いコスト
                                firstTime = False
                            Else
                                ret.money_SONEKI += t.price * e.volume - e.c_money - e.c_cost
                            End If
                            t.exited_volume += e.volume
                            e.exited_volume += e.volume
                            ret.money_exit = t.c_money
                        ElseIf entry_moreThen_exit Then '仕掛け残り枚数が手仕舞い以上
                            ret.money_SONEKI = t.c_money - t.c_cost - (e.price * t.volume) - e.c_cost
                            If 0 < e.exited_volume Then ret.money_SONEKI += e.c_cost '仕掛けコストのダブリ調整
                            e.exited_volume += t.volume
                            ret.money_exit = t.c_money
                            Exit For
                        End If
                    Next
                Case 0 <= t.type_trade.IndexOf("信用返済売")
                    ret.money_enter = ""
                    ret.money_exit = ret.c_money
                    ret.money_SONEKI = ret.money
                    Dim entres = SHINYO_buy_entry(t, retList).Where(Function(s) s.c_money = ret.c_money - ret.money_SONEKI - ret.c_cost).ToList
                    For Each e In entres '処理済みリスト
                        Dim IsExited = e.volume - e.exited_volume < t.volume
                        If IsExited Then Continue For '残出来高が小さければ次へ
                        e.exited_volume += t.volume
                    Next
                Case 0 <= t.type_trade.IndexOf("信用返済買")
                    ret.money_enter = ""
                    ret.money_exit = ret.c_money * -1
                    ret.money_SONEKI = ret.money
                    Dim sells = SHINYO_sell_entry(t, retList)
                    Dim entres = sells.Where(Function(s) -1 * s.c_money = -1 * ret.c_money - ret.money_SONEKI - ret.c_cost).ToList
                    For Each e In entres '処理済みリスト
                        Dim IsExited = e.volume - e.exited_volume < t.volume
                        If IsExited Then Continue For '残出来高が小さければ次へ
                        e.exited_volume += t.volume
                    Next
                Case Else
            End Select
            retList.Add(ret)
        Next
        Return retList
    End Function
    Private Function orderBy(_historys As List(Of history_t)) As List(Of history_t)
        Return _historys.OrderBy(Function(t0) t0.code).ThenBy(Function(t1) t1.date_trade).ThenBy(Function(t1) t1.orderBytype).ToList
    End Function
    Private Function GENBUTSU_entry(_t As history_t, _source As List(Of history_t)) As List(Of history_t)
        Dim entres = _source.Where(Function(s) s.code = _t.code AndAlso s.IsGenbutsuEntry AndAlso Date.Compare(s.date_trade, _t.date_trade) <= 0).ToList
        Return orderBy(entres)
    End Function '現売以前の現物仕掛けを抽出
    Private Function SHINYO_sell_entry(_t As history_t, _source As List(Of history_t)) As List(Of history_t)
        Dim entres = _source.Where(Function(s) s.code = _t.code AndAlso s.IsMeaginSellEntry AndAlso Date.Compare(s.date_trade, _t.date_trade) <= 0).ToList
        Return orderBy(entres)
    End Function '返済以前の信用新規売を抽出
    Private Function SHINYO_buy_entry(_t As history_t, _source As List(Of history_t)) As List(Of history_t)
        Dim entres = _source.Where(Function(s) s.code = _t.code AndAlso s.IsMeaginBuyEntry AndAlso Date.Compare(s.date_trade, _t.date_trade) <= 0).ToList
        Return orderBy(entres)
    End Function '返済以前の信用新規買を抽出
    Private Function copyHistory(_history As history_t) As history_t
        Dim _copy = New history_t
        _copy.AZUKARI = _history.AZUKARI
        _copy.code = _history.code
        _copy.cost = _history.cost
        _copy.date_trade = _history.date_trade
        _copy.KAZEI = _history.KAZEI
        _copy.limit_margin = _history.limit_margin
        _copy.market = _history.market
        _copy.money = _history.money
        _copy.name = _history.name
        _copy.price = _history.price
        _copy.tax = _history.tax
        _copy.type_trade = _history.type_trade
        _copy.UKEWATASHIBI = _history.UKEWATASHIBI
        _copy.volume = _history.volume
        _copy.volumeCount = _history.volumeCount
        Return _copy
    End Function

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
                                            .exited_volume = 0}
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
    Private Sub writeCsv(Of T)(_fname As String, _source As List(Of T))
        Using StreamWriter = New StreamWriter(_fname, False, Encoding.GetEncoding("shift_jis"))
            Using csv = New CsvWriter(StreamWriter)
                csv.Configuration.HasHeaderRecord = True
                csv.WriteRecords(_source)
            End Using
        End Using
    End Sub
End Class
