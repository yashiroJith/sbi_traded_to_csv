Imports System.Text
Imports System.IO
Imports System.Environment
Imports CsvHelper
Imports System.IO.File

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
        textLog.AppendText(fileName(0) & vbCrLf & vbCrLf) 'ソース名を表示
        Dim header = map_header(fileName(0)) 'SBI取引履歴の情報抽出
        Dim query_date = map_traded(fileName(0)) '日付順に出力
        Dim query_code = orderByCode(query_date) '銘柄別に出力
        Dim soneki = Calc_SONEKI((query_date))
        Dim form_soneki = swap_soneki(soneki)
        writeCsv(Path.Combine(CurrentDirectory, "Debug.csv"), soneki)
        'writeCsv(Path.Combine(CurrentDirectory, "Trading_codeSort.csv"), form_soneki)
        writeCsv(Path.Combine(CurrentDirectory, "Trading_dateSort.csv"), form_soneki)

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
                Case 0 <= t.type_trade.IndexOf("信用新規")
                    SHINYO_entry(ret)
'                Case 0 <= t.type_trade.IndexOf("現渡") '信用分を取り消す
'                    ret.money_enter = ret.c_money
'                    ret.money_exit = ""
'                    ret.money_SONEKI = ""
'                    Dim entres = SHINYO_sell_entry(t, retList).Where(Function(s) s.price = ret.price).ToList
'                    For Each e In entres '処理済みリスト
'                        Dim IsExited = e.remainVolume < t.volume
'                        If IsExited Then Continue For '残出来高が小さければ次へ
'                        e.money_enter = e.price * (e.volume - t.volume) '処理済み金額を差し引く
'                        e.exited_volume += t.volume
'                    Next
'                Case 0 <= t.type_trade.IndexOf("株式現物買")
'                    '未手仕舞いの同銘柄現物があれば取得金額、株価の平均を取る
'                    averagePricing(ret, GENBUTSU_entry(t, retList))
'                Case 0 <= t.type_trade.IndexOf("現引") '信用分を取り消す
'                    Dim entres = SHINYO_buy_entry(t, retList).Where(Function(s) s.price = ret.price).ToList '過去の信用新規買リスト
'                    For Each e In entres '該当信用新規買の処理
'                        Dim IsExited = e.remainVolume < t.volume
'                        If IsExited Then Continue For '残出来高が小さければ次へ
'                        e.money_enter = e.price * (e.volume - t.volume) '処理済み金額を差し引く
'                        e.exited_volume += t.volume
'                    Next
'                    'ここから株式現物買と同じ処理
'                    averagePricing(ret, GENBUTSU_entry(t, retList))
'                Case 0 <= t.type_trade.IndexOf("株式現物売") '以前の現引きか現物買いから損益を計算
'                    ret.money_enter = ""
'                    Dim entres = GENBUTSU_entry(t, retList)
'                    Dim firstTime As Boolean = True
'                    For Each e In entres '仕掛けリスト
'#Region "手仕舞い条件"
'                        Dim has_exit_volume = t.exited_volume < t.volume '手仕舞い残りがある
'                        Dim has_entry_volume = 0 < e.remainVolume '仕掛け残りがある
'                        Dim exit_moreThen_entry = e.remainVolume < t.volume '手仕舞いが仕掛け残りより多い
'                        Dim entry_moreThen_exit = t.volume <= e.remainVolume '仕掛け残り枚数が手仕舞い以上
'#End Region
'                        If exit_moreThen_entry AndAlso has_entry_volume AndAlso has_exit_volume Then '手仕舞いが仕掛けより多い
'                            If firstTime Then
'                                ret.money_SONEKI += t.price * e.volume - t.c_cost - e.c_money  '最初だけ手仕舞いコスト
'                                ret.cost_exit = t.c_cost
'                                ret.date_enter = e.UKEWATASHIBI
'                                firstTime = False
'                            Else
'                                ret.money_SONEKI += t.price * e.volume - e.c_money
'                            End If
'                            t.exited_volume += e.volume
'                            e.exited_volume += e.volume
'                            ret.money_exit = t.c_money - t.c_cost
'                        ElseIf entry_moreThen_exit Then '仕掛け残り枚数が手仕舞い以上 一回で処理が終了する時
'                            ret.money_SONEKI = t.c_money - t.c_cost - (e.price * t.volume)
'                            e.exited_volume += t.volume
'                            ret.money_exit += t.c_money - t.c_cost
'                            ret.cost_exit = t.c_cost
'                            ret.date_enter = e.UKEWATASHIBI
'                            Exit For
'                        End If
'                    Next

                Case 0 <= t.type_trade.IndexOf("信用返済売")
                    ret.money_SONEKI = ret.money
                    Dim entres = SHINYO_buy_entry(ret, retList).ToList
                    For Each e In entres 'エントリー抽出リスト
                        Dim entry_price = (ret.c_money - ret.c_cost - ret.money_SONEKI) / ret.volume
                        If e.remainVolume < t.volume Then Continue For '手仕舞い株数より仕掛け株数が少ないときだけ進む
                        If e.price = entry_price Then 'エントリーからの仕掛け値と手仕舞いから計算した仕掛け値が一致
                            If e.volume = t.volume Then '仕掛けと返済数が同じ時
                                'ret <- e 仕掛けから手仕舞いにデータをコピー
                                ret.money_enter_cost = e.money_enter_cost
                                ret.money_enter = e.money_enter
                                ret.cost_enter = e.cost_enter
                                ret.cost_exit = ret.c_cost - e.cost_enter
                                ret.money_exit = ret.c_money
                                ret.money_exit_cost = ret.c_money - ret.cost_exit
                            Else '仕掛けが多く分割返済の時
                                ret.cost_enter = Math.Round(e.c_cost / e.volume * ret.volume, 0, MidpointRounding.AwayFromZero)
                                ret.cost_exit = ret.c_cost - ret.cost_enter
                                ret.money_enter = e.price * ret.volume
                                ret.money_enter_cost = ret.money_enter + ret.cost_enter
                                ret.money_exit = ret.c_money
                                ret.money_exit_cost = ret.c_money - ret.cost_exit
                            End If
                            e.exited_volume += t.volume
                        End If

                    Next
                    'ret.money_exit = ret.c_money
                    'ret.money_exit_cost = ret.c_money - t.cost_exit

                    'Case 0 <= t.type_trade.IndexOf("信用返済買")
                    '    ret.money_SONEKI = ret.money
                    '    'Dim entres = SHINYO_sell_entry(t, _source).Where(Function(s) ret.money_SONEKI = s.c_money - t.c_money - ret.c_cost).ToList
                    '    Dim entres = SHINYO_sell_entry(t, _source)
                    '    For Each e In entres 'エントリー抽出リスト
                    '        Dim t_money = t.c_money
                    '        Dim ret_money = ret.c_money
                    '        Dim e_money = e.c_money
                    '        Dim _soneki = t.c_money - e.c_money
                    '        ret.exited_volume += t.volume
                    '        ret.money_enter = e.money_enter '仕掛け金額を手仕舞いに累算
                    '        ret.cost_enter = e.c_cost '仕掛けコストを手仕舞いに累算
                    '        If ret.date_enter <> "" Then Continue For
                    '        ret.date_enter = e.date_enter
                    '        ret.cost_exit = t.c_cost - ret.cost_enter
                    '    Next
                    '    ret.money_exit = ret.c_money - ret.cost_exit
                    'Case 0 <= t.type_trade.IndexOf("配当金")
                    '    ret.money_enter = ""
                    '    ret.money_exit = ""
                    '    ret.money_SONEKI = ret.money
                Case Else
            End Select
            retList.Add(ret)
        Next
        Return retList
    End Function
    Private Sub SHINYO_entry(ByRef ret As history_t)
        ret.money_enter = ret.c_money '仕掛け金額
        ret.cost_enter = ret.c_cost
        ret.date_enter = ret.UKEWATASHIBI
        'ret.cost_enter_unit = Math.Floor(ret.c_cost / ret.volume * 100)
        If ret.IsMeaginBuyEntry Then
            ret.money_enter_cost = ret.c_money + ret.c_cost '買いエントリー
        Else
            ret.money_enter_cost = ret.c_money - ret.c_cost '売りエントリー
        End If
    End Sub '信用新規 -> 仕掛け金額 -> 仕掛け分コスト計算
    Private Sub matching_SHINYO_entry_exit(ByRef entres As List(Of history_t), ByRef ret As history_t, ByRef t As history_t)
        For Each e In entres 'エントリー抽出リスト
            Dim t_money = t.c_money
            Dim ret_money = ret.c_money
            Dim e_money = e.c_money
            Dim _soneki = If(e.type_trade.IndexOf("買"), t.c_money - e.c_money, e.c_money - t.c_money)
            If ret.money_SONEKI = _soneki - ret.c_cost Then '1取引で損益の合うもの
                e.exited_volume = t.volume
                ret.money_enter = e.money_enter '仕掛け金額を手仕舞いに累算
                ret.cost_enter = e.c_cost '仕掛けコストを手仕舞いに累算
                If ret.date_enter <> "" Then Continue For
                ret.date_enter = e.date_enter
                ret.cost_exit = t.c_cost - ret.cost_enter
            ElseIf t.volume < e.remainVolume Then
                e.exited_volume += t.volume
                ret.cost_enter = Math.Floor(e.c_cost / e.volume * t.volume)
                ret.money_enter += e.money_enter + e.cost_enter
                ret.cost_exit = t.c_cost - ret.cost_enter
                ret.date_enter = e.date_enter
            End If
            'If ret.money_SONEKI = _soneki - ret.c_cost Then '1取引で損益の合うもの
            '    e.exited_volume = t.volume
            '    ret.money_enter = e.money_enter '仕掛け金額を手仕舞いに累算
            '    ret.cost_enter = e.c_cost '仕掛けコストを手仕舞いに累算
            '    If ret.date_enter <> "" Then Continue For
            '    ret.date_enter = e.date_enter
            '    ret.cost_exit = t.c_cost - ret.cost_enter
            'ElseIf t.volume < e.remainVolume AndAlso ret.cost_exit <> "" Then
            '    e.exited_volume += t.volume
            '    ret.money_enter += e.money_enter
            '    ret.cost_enter = Math.Floor(e.c_cost / e.volume * t.volume)
            '    ret.date_enter = e.date_enter
            '    ret.cost_exit = t.c_cost - ret.cost_enter
            'End If
        Next
        ret.money_exit = ret.c_money - ret.cost_exit
    End Sub
    Private Sub averagePricing(ByRef ret As history_t, entres As List(Of history_t))
        Dim remainMoney As Double = 0
        Dim remainVolume As Integer = 0
        For Each e In entres
            If 0 < e.remainVolume Then '処理待ちの株があれば購入金額を集計
                remainMoney += e.remainVolume * e.price + e.c_cost
                remainVolume += e.remainVolume
            End If
        Next
        '過去購入金額が0なら新規株価、そうでなければ平均法
        If remainMoney = 0 Then
            ret.price = Math.Ceiling((ret.c_money + ret.c_cost) / ret.volume)
            ret.money_enter = ret.c_money * -1
        Else
            remainMoney += ret.c_money + ret.c_cost
            remainVolume += ret.volume
            ret.price = Math.Ceiling(remainMoney / remainVolume)
            ret.money_enter = ret.c_money * -1
        End If
        For Each e In entres
            If 0 < e.remainVolume Then '処理待ちの株があれば購入株価と金額を編集
                e.price = ret.price
                e.money_enter = e.c_money * -1
            End If
        Next
        ret.cost = ""
        ret.tax = ""
        ret.money_exit = ""
        ret.money_SONEKI = ""
    End Sub '口座を調べてすでに同じ銘柄の現物があれば平均を取る。小数点は切り上げ。
    Private Function orderByDate(_historys As List(Of history_t)) As List(Of history_t)
        Return _historys.OrderBy(Function(t1) t1.date_trade).ThenBy(Function(t1) t1.orderBytype).ToList
    End Function '日付順
    Private Function orderByCode(_historys As List(Of history_t)) As List(Of history_t)
        Return _historys.OrderBy(Function(t0) t0.code).ThenBy(Function(t1) t1.date_trade).ThenBy(Function(t1) t1.orderBytype).ToList
    End Function '視覚的に見やすい順
    Private Function GENBUTSU_entry(_t As history_t, _source As List(Of history_t)) As List(Of history_t)
        Dim entres = _source.Where(Function(s) (s.code = _t.code OrElse s.code = _t.code & "1") AndAlso s.IsGenbutsuEntry AndAlso Date.Compare(s.date_trade, _t.date_trade) <= 0).ToList
        Return orderByCode(entres)
    End Function '現売以前の現物仕掛けを抽出
    Private Function SHINYO_sell_entry(_t As history_t, _source As List(Of history_t)) As List(Of history_t)
        Dim entres = _source.Where(Function(s) s.code = _t.code AndAlso s.IsMeaginSellEntry AndAlso Date.Compare(s.date_trade, _t.date_trade) <= 0).ToList
        Return orderByCode(entres)
    End Function '返済以前の信用新規売を抽出
    Private Function SHINYO_buy_entry(_t As history_t, _source As List(Of history_t)) As List(Of history_t)
        Dim entres = _source.Where(Function(s) s.code = _t.code AndAlso s.IsMeaginBuyEntry AndAlso Date.Compare(s.date_trade, _t.date_trade) <= 0).ToList
        Return orderByCode(entres)
    End Function '返済以前の信用新規買を抽出
    Private Function SHINYO_sell_exit(_t As history_t, _source As List(Of history_t)) As List(Of history_t)
        Dim exits = _source.Where(Function(s) s.code = _t.code AndAlso s.IsMeaginSellExit AndAlso Date.Compare(s.date_trade, _t.date_trade) <= 0).ToList
        Return orderByCode(exits)
    End Function
    Private Function copyHistory(_history As history_t) As history_t
        Dim filter = Function(_str As String) As String
                         Return If(_str = "--", "0", _str)
                     End Function
        Dim _copy = New history_t
        _copy.AZUKARI = _history.AZUKARI
        _copy.code = _history.code
        _copy.cost = filter(_history.cost)
        _copy.date_trade = _history.date_trade
        _copy.KAZEI = _history.KAZEI
        _copy.limit_margin = _history.limit_margin
        _copy.market = _history.market
        _copy.money = filter(_history.money)
        _copy.name = _history.name
        _copy.price_source = filter(_history.price)
        _copy.price = filter(_history.price)
        _copy.tax = filter(_history.tax)
        _copy.type_trade = _history.type_trade
        _copy.UKEWATASHIBI = _history.UKEWATASHIBI
        _copy.volume = filter(_history.volume)
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
                                            .price_source = _index(9),
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
        Dim _header As New header_t
        Try
            _header = New header_t With {.SYOUHIN_SHITEI = _headed(0),
                                         .KAISHI_DATE = CDate(_headed(1)).ToShortDateString,
                                         .SYURYOU_DATE = CDate(_headed(2)).ToShortDateString,
                                         .MEISAI_SUU = _headed(3),
                                         .MEISAI_KAISSHI = _headed(4),
                                         .MEISAI_SYURYOU = _headed(5)}

        Catch ex As Exception
            textLog.AppendText($"このファイルはSBI取引履歴ではありません" & vbCrLf & vbCrLf)
        End Try
        Return _header
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
        Dim flag As Boolean = False
        Do
            Try
                Using StreamWriter = New StreamWriter(_fname, False, Encoding.GetEncoding("shift_jis"))
                    Using csv = New CsvWriter(StreamWriter)
                        csv.Configuration.HasHeaderRecord = True
                        csv.WriteRecords(_source)
                    End Using
                End Using
                flag = True
            Catch ex As Exception
                MsgBox($"{_fname}が開かれています。")
            End Try
        Loop Until flag
    End Sub
End Class
