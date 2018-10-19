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
                            Text = "SBI一般口座支援ツール"
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
        Dim history = map_history(fileName(0), 10) '10行目から
        Dim patch_record = map_history(Path.Combine(CurrentDirectory, "patch.csv"), 2) '2行目から

        do_patch(history, patch_record) 'パッチデータを適用
        Dim all_history = calculate(orderByDate(history)) '日付順に出力

        '分割で自動出力できなかったものを出力
        Dim pre_illegal = all_history.Where(Function(r) (0 <= r.TORIHIKI.IndexOf("信用返済") AndAlso r.money_enter_ct = 0) OrElse (0 <= r.TORIHIKI.IndexOf("現物売") AndAlso 0 < r.remainVolume))
        Dim illegal = all_history.Where(Function(r) pre_illegal.Any(Function(i) i.code = r.code AndAlso r.exted_vol = 0)).ToList

        writeCsv(Path.Combine(CurrentDirectory, "Debug.csv"), swap_history_to_patch(orderByCode(all_history)))
        writeCsv(Path.Combine(CurrentDirectory, "Result.csv"), swap_history_to_MEISAI(all_history))
        writeCsv(Path.Combine(CurrentDirectory, "Illegal.csv"), swap_history_to_patch(illegal))


    End Sub 'ファイルをドラッグ＆ドロップする処理
    Private Sub do_patch(ByRef s As List(Of history_t), _patch As List(Of history_t))
        For i As Integer = s.Count = 1 To 0 Step -1
            For Each p In _patch
                If s(i).code = p.code AndAlso CDate(s(i).dt_YAKUJYOU) = CDate(p.dt_YAKUJYOU) AndAlso s(i).TORIHIKI = p.TORIHIKI AndAlso s(i).cost_sr = p.cost_sr Then
                    s.RemoveAt(i)
                End If
            Next
        Next
        For Each p In _patch
            s.Add(p)
        Next
    End Sub

    Private Function swap_history_to_MEISAI(exits As List(Of history_t)) As List(Of SONEKI_MEISAI_t)
        Dim trades = New List(Of SONEKI_MEISAI_t)
        For Each e In exits
            If (0 <= e.TORIHIKI.IndexOf("現物売")) OrElse (0 <= e.TORIHIKI.IndexOf("信用返済")) OrElse 0 <= e.TORIHIKI.IndexOf("配当金") Then
                Dim trade = New SONEKI_MEISAI_t
                trade.銘柄コード = e.code
                trade.銘柄 = e.name
                trade.譲渡益取消 = ""
                trade.約定日 = e.dt_YAKUJYOU
                trade.数量 = e.volume
                trade.取引 = e.TORIHIKI
                trade.受渡日 = e.UKEWATASHIBI
                trade.売却l決済金額 = e.money_exit_ct
                trade.費用 = e.cost_exit
                trade.取得l新規年月日 = e.dt_enter
                trade.取得l新規金額 = e.money_enter_ct
                trade.損益金額 = e.SONEKI
                trades.Add(trade)
            End If
        Next
        Return trades
    End Function
    Private Function swap_history_to_patch(exits As List(Of history_t)) As List(Of patch_t)
        Dim patchs = New List(Of patch_t)
        For Each e In exits
            Dim patch = New patch_t
            patch.約定日 = e.dt_YAKUJYOU
            patch.銘柄 = e.name
            patch.銘柄コード = e.code
            patch.市場 = e.market
            patch.取引 = e.TORIHIKI
            patch.期限 = e.KIGEN
            patch.預り = e.AZUKARI
            patch.課税 = e.KAZEI
            patch.約定数量 = e.volume
            patch.約定単価 = e.TANKA_sr
            patch.手数料 = e.cost_sr
            patch.税額 = e.tax_sr
            patch.受渡日 = e.UKEWATASHIBI
            patch.受渡金額 = e.money_sr
            patch.処理済数量 = e.exted_vol
            patchs.Add(patch)
        Next
        Return patchs
    End Function

    Private Function calculate(_source As List(Of history_t)) As List(Of history_t)
        Dim retList = New List(Of history_t)
        For Each t In _source
            Dim ret = New history_t
            ret = copyHistory(t)
            Select Case True
                Case 0 <= ret.TORIHIKI.IndexOf("信用新規") '信用新規 -> 仕掛け金額 -> 仕掛け分コスト計算
                    ret.money_enter = ret.c_money_shinyo '仕掛け金額
                    ret.cost_enter = ret.c_cost '仕掛けコスト
                    ret.dt_enter = ret.UKEWATASHIBI
                    If ret.IsMeaginBuyEntry Then
                        ret.money_enter_ct = ret.c_money_shinyo + ret.c_cost 'コスト入り新規買い金額
                    Else
                        ret.money_enter_ct = ret.c_money_shinyo - ret.c_cost 'コスト入り新規売り金額
                    End If
                Case 0 <= ret.TORIHIKI.IndexOf("現渡") '信用売りを現物から取り消す
                    Dim ShinyoSell = SHINYO_sell_entry(ret, retList).Where(Function(e) e.TANKA_sr = ret.TANKA_sr).ToList '同単価
                    For Each s In ShinyoSell '信用売りリスト
                        If s.remainVolume < ret.volume Then Continue For '残出来高が小さければ次へ
                        s.exted_vol += ret.volume '信用売側を処理済み
                        Exit For '現渡しのルールから１度当たれば終了
                    Next
                    Dim entres = GENBUTSU_entry(ret, retList)
                    Dim Genbutu = New history_t
                    For Each e In entres
                        If e.remainVolume < ret.volume Then Continue For '残出来高が小さければ次へ
                        e.exted_vol += ret.volume
                        Genbutu = e
                        Exit For '現渡しのルールから１度当たれば終了
                    Next
                    ret.dt_enter = Genbutu.UKEWATASHIBI
                    ret.money_enter_ct = Genbutu.money_enter_ct
                    ret.money_exit_ct = ret.c_money_shinyo - ret.c_cost '売却・決済金額
                    ret.cost_exit = ret.c_cost '売却費用
                    ret.SONEKI = ret.money_exit_ct - ret.money_enter_ct
                    '現物から取得分 ->　コスト入り取得金額 - 損益
                Case 0 <= ret.TORIHIKI.IndexOf("株式現物買") OrElse 0 <= ret.TORIHIKI.IndexOf("分売買")
                    averagePricing(ret, GENBUTSU_entry(ret, retList))
                Case 0 <= ret.TORIHIKI.IndexOf("現引") '信用分を取り消す
                    Dim entres = SHINYO_buy_entry(t, retList).Where(Function(s) s.TANKA_ct = ret.TANKA_ct).ToList '過去の信用新規買リスト
                    Select Case True
                        Case entres.Count = 0
                            Continue For
                        Case entres.Count = 1
                            Dim e = entres(0)
                            ret.dt_enter = e.UKEWATASHIBI '現引 -> 受渡日
                            e.exted_vol += t.volume '信用 -> exted_vol
                        Case 1 < entres.Count
                            For Each e In entres '複数新規買の処理
                                If e.remainVolume < ret.volume Then Continue For '残出来高が小さければ次へ
                                If ret.TANKA_sr = e.TANKA_sr Then
                                    ret.dt_enter = e.UKEWATASHIBI '現引 -> 受渡日
                                    e.exted_vol += ret.volume '信用 -> exted_vol
                                End If
                            Next
                    End Select
                    ret.money_sr *= -1
                    'ここから株式現物買と同じ処理
                    averagePricing(ret, GENBUTSU_entry(ret, retList))
                Case 0 <= ret.TORIHIKI.IndexOf("株式現物売")  '以前の現引きか現物買いから損益を計算
                    'entresは複数だと買い増しした現物、もしくは単独。
                    Dim entres = GENBUTSU_entry(ret, retList).Where(Function(e) e.exted_vol <> e.volume).ToList
                    Select Case True
                        Case ret.KAZEI = "非課税"
                            Continue For'NISAは無視
                        Case entres.Count = 0 '仕掛けが０
                            If 0 <= ret.name.IndexOf("新株予約権") Then
                                ret.cost_exit = ret.c_cost
                                ret.dt_enter = ret.dt_YAKUJYOU
                                ret.money_exit_ct = ret.money_sr
                                ret.SONEKI = ret.money_sr
                            End If
                        Case entres.Count = 1 '仕掛けが単独
                            Dim e = entres(0)
                            Dim min_vol = Math.Min(e.volume, ret.volume) '小さい株数が処理済み
                            ret.exted_vol += min_vol '小さい株数が処理済み
                            e.exted_vol += min_vol  '小さい株数が処理済み
                            ret.money_enter_ct = e.TANKA_ct * ret.volume '税金上の取得金額
                            ret.SONEKI = CInt(ret.money_sr) - ret.money_enter_ct
                            ret.money_exit_ct = ret.money_sr '表示用コスト込金額に代入
                            ret.cost_exit = ret.c_cost '
                            ret.dt_enter = e.UKEWATASHIBI
                        Case 1 < entres.Count '仕掛けが複数 -> 買い増しされた仕掛け
                            For Each e In entres
                                Dim min_vol = Math.Min(e.remainVolume, ret.remainVolume) '小さい株数が処理済み
                                ret.exted_vol += min_vol
                                e.exted_vol += min_vol
                                ret.money_enter_ct = e.TANKA_ct * ret.volume '税金上の取得金額
                                ret.SONEKI = CInt(ret.money_sr) - ret.money_enter_ct '損益計算
                                ret.money_exit_ct = ret.money_sr '表示用コスト込金額に代入
                                ret.cost_exit = ret.c_cost '
                                ret.dt_enter = e.UKEWATASHIBI
                                e.money_enter_ct = e.TANKA_ct * min_vol
                                If ret.exted_vol = ret.volume Then Exit For '手仕舞い株式の処理が終わり
                            Next
                    End Select
                Case 0 <= ret.TORIHIKI.IndexOf("信用返済売")
                    ret.SONEKI = ret.money_sr
                    Dim entres = SHINYO_buy_entry(ret, retList).ToList
                    matching_SHINYO_entry_exit(entres, ret)
                Case 0 <= ret.TORIHIKI.IndexOf("信用返済買")
                    ret.SONEKI = ret.money_sr
                    Dim entres = SHINYO_sell_entry(ret, retList) '売り買い逆
                    matching_SHINYO_entry_exit(entres, ret)
                Case 0 <= t.TORIHIKI.IndexOf("配当金")
                    ret.money_enter = 0
                    ret.money_exit = 0
                    ret.SONEKI = ret.money_sr
                Case Else
            End Select
            retList.Add(ret)
        Next
        Return retList
    End Function
    Private Sub averagePricing(ByRef ret As history_t, entres As List(Of history_t))
        '未手仕舞いの同銘柄現物があれば取得金額、株価の平均を取る
        Dim add_entres = entres.Where(Function(e) 0 < e.remainVolume)
        If 0 = add_entres.Count Then '現物買いが新規買い
            ret.TANKA_ct = Math.Ceiling(ret.money_sr / ret.volume) '単価(ｺｽﾄ込)= 受渡金額(ｺｽﾄ込金額)/数量
            ret.money_enter_ct = ret.TANKA_ct * ret.volume '取得金額 ok
        ElseIf 0 < add_entres.Count Then '既存株がある
            Dim Moneys As Double = ret.money_sr
            Dim Volumes As Integer = ret.volume
            For Each e In add_entres
                Moneys += e.money_sr  '購入済み金額（コスト入り）
                Volumes += e.volume '購入済み株数
            Next
            ret.TANKA_ct = Math.Ceiling(Moneys / Volumes) '税金上の取得単価 -> 取得金額は 売却側で計算
            For Each e In add_entres
                e.TANKA_ct = ret.TANKA_ct '既存株の税金上の取得単価 -> 取得金額は 売却側で計算
            Next
        End If
    End Sub '口座を調べてすでに同じ銘柄の現物があれば平均を取る。小数点は切り上げ。
    Private Sub matching_SHINYO_entry_exit(ByRef entres As List(Of history_t), ByRef ret As history_t)
        Dim vector = If(0 <= ret.TORIHIKI.IndexOf("信用返済買"), 1, -1)
        For Each e In entres 'エントリー抽出リスト
            If e.code = "3475" Then
                Dim a = 0
            End If
            Dim entry_price = (ret.c_money_shinyo + (ret.c_cost + ret.SONEKI) * vector) / ret.volume '売り買い逆
            If e.remainVolume < ret.volume Then Continue For '手仕舞い株数より仕掛け株数が少ないときだけ進む
            If e.TANKA_sr = entry_price Then 'エントリーからの仕掛け値と手仕舞いから計算した仕掛け値が一致
                If e.volume = ret.volume Then '仕掛けと返済数が同じ時
                    ret.money_enter_ct = e.money_enter_ct 'コスト入り新規金額
                    ret.cost_enter = e.cost_enter '新規コスト
                    ret.cost_exit = ret.c_cost - e.cost_enter '手仕舞いコスト
                    ret.money_exit_ct = ret.c_money_shinyo + (ret.money_exit + ret.cost_exit) * vector 'コスト入り手仕舞い金額 -> 売り買い逆
                Else '仕掛けが多く分割返済の時
                    ret.cost_enter = Math.Round(e.c_cost / e.volume * ret.volume, 0, MidpointRounding.AwayFromZero) '新規コスト
                    ret.cost_exit = ret.c_cost - ret.cost_enter '手仕舞いコスト
                    ret.money_enter_ct = e.TANKA_sr * ret.volume + ret.cost_enter 'コスト入り新規金額
                    ret.money_exit_ct = ret.c_money_shinyo + (ret.money_exit + ret.cost_exit) * vector 'コスト入り手仕舞い金額 -> 売り買い逆
                End If
                ret.dt_enter = e.UKEWATASHIBI
                e.exted_vol += ret.volume
                ret.exted_vol += Math.Min(ret.volume, e.volume)
            End If
        Next
    End Sub
    Private Function orderByDate(_historys As List(Of history_t)) As List(Of history_t)
        Return _historys.OrderBy(Function(t1) CDate(t1.dt_YAKUJYOU)).ThenBy(Function(t1) t1.orderBytype).ToList
    End Function '日付順
    Private Function orderByCode(_historys As List(Of history_t)) As List(Of history_t)
        Return _historys.OrderBy(Function(t0) t0.code).ThenBy(Function(t1) t1.dt_YAKUJYOU).ThenBy(Function(t1) t1.orderBytype).ToList
    End Function '視覚的に見やすい順
    Private Function GENBUTSU_entry(_t As history_t, _source As List(Of history_t)) As List(Of history_t)
        Dim entres = _source.Where(Function(s) (s.code = _t.code OrElse s.code = _t.code & "1") AndAlso s.IsGenbutsuEntry AndAlso Date.Compare(s.dt_YAKUJYOU, _t.dt_YAKUJYOU) <= 0).ToList
        Return entres
    End Function '現売以前の現物仕掛けを抽出
    Private Function SHINYO_sell_entry(_ret As history_t, _retList As List(Of history_t)) As List(Of history_t)
        Dim entres = _retList.Where(Function(r) r.code = _ret.code AndAlso r.IsMeaginSellEntry AndAlso Date.Compare(r.dt_YAKUJYOU, _ret.dt_YAKUJYOU) <= 0).ToList
        Return entres
    End Function '返済以前の信用新規売を抽出
    Private Function SHINYO_buy_entry(_t As history_t, _source As List(Of history_t)) As List(Of history_t)
        Dim entres = _source.Where(Function(s) s.code = _t.code AndAlso s.IsMeaginBuyEntry AndAlso Date.Compare(s.dt_YAKUJYOU, _t.dt_YAKUJYOU) <= 0).ToList
        Return entres
    End Function '返済以前の信用新規買を抽出
    Private Function SHINYO_sell_exit(_t As history_t, _source As List(Of history_t)) As List(Of history_t)
        Dim exits = _source.Where(Function(s) s.code = _t.code AndAlso s.IsMeaginSellExit AndAlso Date.Compare(s.dt_YAKUJYOU, _t.dt_YAKUJYOU) <= 0).ToList
        Return exits
    End Function
    Private Function copyHistory(_history As history_t) As history_t
        Dim filter = Function(_str As String) As String
                         Return If(_str = "--", "0", _str)
                     End Function
        Dim _copy = New history_t
        _copy.AZUKARI = _history.AZUKARI
        _copy.code = _history.code
        _copy.dt_YAKUJYOU = _history.dt_YAKUJYOU
        _copy.KAZEI = _history.KAZEI
        _copy.KIGEN = _history.KIGEN
        _copy.market = _history.market
        _copy.money_sr = filter(_history.money_sr)
        _copy.name = _history.name
        _copy.TANKA_sr = filter(_history.TANKA_sr)
        _copy.cost_sr = filter(_history.cost_sr)
        _copy.tax_sr = filter(_history.tax_sr)
        _copy.TORIHIKI = _history.TORIHIKI
        _copy.UKEWATASHIBI = _history.UKEWATASHIBI
        _copy.volume = filter(_history.volume)
        Return _copy
    End Function
#End Region
    Private Function map_history(_fname As String, _startLine As Integer) As List(Of history_t)
        Dim result = read_from_csv(_fname, _startLine)
        Dim csv = result.Split(vbCrLf)
        Dim lines As New List(Of history_t)
        Dim _buffer As String = ""
        For i As Integer = 0 To csv.Count - 1
            _buffer = csv(i).Replace(vbLf, "")
            Dim _index() = _buffer.Split(",")
            If _buffer = "" Then Continue For
            Dim traded = New history_t With {.dt_YAKUJYOU = _index(0),'約定日
                                            .name = _index(1),'銘柄名
                                            .code = _index(2),'銘柄コード
                                            .market = _index(3),'
                                            .TORIHIKI = _index(4),
                                            .KIGEN = _index(5),
                                            .AZUKARI = _index(6),
                                            .KAZEI = _index(7),
                                            .volume = _index(8),
                                            .TANKA_sr = _index(9),
                                            .cost_sr = _index(10),
                                            .tax_sr = _index(11),
                                            .UKEWATASHIBI = _index(12),
                                            .money_sr = _index(13),'受渡金額
                                            .exted_vol = 0}
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
