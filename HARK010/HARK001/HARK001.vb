'/*-----------------------------------------------------------------------------
' * COPYRIGHT(C) ITI CORPORATION 2020
' * ITI CONFIDENTIAL AND PROPRIETARY
' *
' * All rights reserved by ITI Corporation.
' *-----------------------------------------------------------------------------/
Option Compare Binary
Option Explicit On
Option Strict On

Imports AdvanceSoftware.VBReport8
Imports HARK010.HARK_DBCommon
Imports HARK010.HARK_Documents
Imports HARK010.HARK_Sub
Imports HARK010.HARK_Common
Imports NAppUpdate.Framework
Imports System.ComponentModel


Public Class HARK001

    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　フォーム読み込み処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　sender・・オブジェクト識別クラス
    ' *　引数２　　　　　：　e・・イベントデータクラス
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Fm_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try
            'DataGridViewちらつき防止
            Dim myType As Type = GetType(DataGridView)
            Dim myPropertyInfo As Reflection.PropertyInfo = myType.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
            myPropertyInfo.SetValue(Dgv, True, Nothing)

            Dgv.AutoResizeColumns()
            'Dgv.DefaultCellStyle.NullValue = "-"

            Text = Set_FormTitle([GetType]().Name, My.Application.Info.Title)
            TSSVersion.Text = "Ver : " & Application.ProductVersion

            'サイドバーコントロール位置調整
            Container検索.Height = Dgv.Height - 6
            Bt検索.Location = New Point(12, Dgv.Height - 70)
            Btクリア.Location = New Point(128, Dgv.Height - 70)
            Bt印刷.Location = New Point(12, Dgv.Height - 37)
            Lbl注意１.Location = New Point(16, Dgv.Height - 100)
            Lbl注意２.Location = New Point(102, Dgv.Height - 100)

            Update_Check()

            'コンボに値設定
            Set_CmbValue()

            'コンポーネント初期化
            Initialize()

            '解像度1024×768以下は起動不可
            If Screen.PrimaryScreen.Bounds.Width < 1024 Or Screen.PrimaryScreen.Bounds.Height < 768 Then
                MsgBox(MSG_COM016 & vbCr & MSG_COM017, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                Exit Sub
            End If

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　「×」ボタンクリック
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　sender -- オブジェクト識別クラス
    ' *　引数２　　　　　：　e      -- イベントデータクラス
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Fm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        Try

            gintMsg = MsgBox(MSG_COM003, CType(MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Question, MsgBoxStyle), My.Application.Info.Title)

            If gintMsg = vbNo Then
                e.Cancel = True
                Exit Sub
            End If

            'メモリ開放
            GC.Collect()

            'Oracle切断
            OraDisConnect()

            Dispose()

            'アプリ終了
            Application.Exit()

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　画面アクティブ時処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　sender・・オブジェクト識別クラス
    ' *　引数２　　　　　：　e・・イベントデータクラス
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Fm_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        Try
            '初期フォーカスを項目に設定
            ActiveControl = Cmb汎用

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　アップデートチェック処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　なし
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Update_Check()

        Try

            log.Info("Update_Check")
            log.Info("現在のバージョン：" & My.Application.Info.Version.ToString)
            log.Info("現在の状態：" & UpdateManager.Instance.State.ToString)

            Select Case (UpdateManager.Instance.State)

                Case UpdateManager.UpdateProcessState.NotChecked

                    HARK_Update.Update_Check()

                Case UpdateManager.UpdateProcessState.Checked

                    HARK_Update.Update_Prepare()

                Case UpdateManager.UpdateProcessState.Prepared

                    HARK_Update.Update_Install()

                Case UpdateManager.UpdateProcessState.AppliedSuccessfully

                    MsgBox(MSG_UPD003, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)

                Case UpdateManager.UpdateProcessState.RollbackRequired

                    HARK_Update.Update_Rollback()

                Case UpdateManager.UpdateProcessState.AfterRestart

                    HARK_Update.Update_CheckAfterRestart()

            End Select

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　コンボボックスに値をセット
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　なし
    ' *　戻値　　　　　　：　なし
    ' *
    ' *-----------------------------------------------------------------------------/
    Private Sub Set_CmbValue()

        Dim i As Integer

        Try

            '事業所一覧
            For i = 0 To gint事業所Cnt - 1

                Cmb事業所.Items.Add(New 事業所一覧(事業所Array(i).str事業所名, 事業所Array(i).int事業所コード))

            Next

            '空白行追加
            Cmb事業所.Items.Add(New 事業所一覧("", 0))


            'サブプログラム一覧
            If My.Settings.事業所コード <> 0 Then

                If DTNP0000_PROC0001(gintSQLCODE, gstrSQLERRM) = False Then
                    MsgBox(gintSQLCODE & "-" & gstrSQLERRM & vbCr & MSG_COM902 & vbCr & MSG_COM901, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                    log.Error(Set_ErrMSG(gintSQLCODE, gstrSQLERRM))
                    Exit Sub
                End If

                For i = 0 To gintサブプログラムCnt - 1

                    Cmb汎用.Items.Add(New サブプログラム一覧(サブプログラムArray(i).strサブプログラム名, サブプログラムArray(i).strサブプログラムコード))

                Next

                '空白行追加
                Cmb汎用.Items.Add(New サブプログラム一覧("", ""))

            End If

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　各種コンポーネント初期化
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　なし
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Initialize()

        Try

            '事業所コンボ制御
            If My.Settings.事業所コード = 0 Then
                xxxint事業所コード = 0
                Cmb事業所.Enabled = True
                Cmb事業所.SelectedIndex = gint事業所Cnt
            Else
                xxxint事業所コード = My.Settings.事業所コード
                Cmb事業所.Enabled = False
                Cmb事業所.Text = gstr部門名
            End If

            Cmb汎用.Text = ""
            xxxstrProgram_ID = ""

            InitializeDetail()

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　各種コンポーネント初期化
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　なし
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub InitializeDetail()

        Try

            'データグリッドビュー初期化
            Dgv.DataSource = Nothing
            TSSRowCount.Text = MSG_COM021 & "0"

            '検索条件初期化
            Cmb得意先.Text = ""
            Cmb得意先.Items.Clear()
            Cmb得意先.Enabled = True
            Cmb需要先.Text = ""
            Cmb需要先.Items.Clear()
            Cmb需要先.Enabled = True
            txtDate.Text = ""
            txtDate.Enabled = True
            txt商品コード.Text = ""
            txt商品コード.AlternateText.DisplayNull.Text = ""
            txt商品コード.AlternateText.DisplayNull.ForeColor = Color.Gray
            txt商品コード.Enabled = True
            txt相手先品番.Text = ""
            txt相手先品番.AlternateText.DisplayNull.Text = ""
            txt相手先品番.AlternateText.DisplayNull.ForeColor = Color.Gray
            txt相手先品番.Enabled = True
            txt得意先.Text = ""
            txt得意先.AlternateText.DisplayNull.Text = ""
            txt得意先.AlternateText.DisplayNull.ForeColor = Color.Gray
            txt得意先.Enabled = True
            txt需要先.Text = ""
            txt需要先.AlternateText.DisplayNull.Text = ""
            txt需要先.AlternateText.DisplayNull.ForeColor = Color.Gray
            txt需要先.Enabled = True


            Lbl得意先.Font = New Font(Lbl得意先.Font, System.Drawing.FontStyle.Regular)
            Lbl得意先.ForeColor = Color.Black
            Lbl需要先.Font = New Font(Lbl需要先.Font, System.Drawing.FontStyle.Regular)
            Lbl需要先.ForeColor = Color.Black
            Lbl対象日.Font = New Font(Lbl対象日.Font, System.Drawing.FontStyle.Regular)
            Lbl対象日.ForeColor = Color.Black
            Lbl商品.Font = New Font(Lbl商品.Font, System.Drawing.FontStyle.Regular)
            Lbl商品.ForeColor = Color.Black
            Lbl相手先品番.Font = New Font(Lbl相手先品番.Font, System.Drawing.FontStyle.Regular)
            Lbl相手先品番.ForeColor = Color.Black
            Lbl得意先コード指定.Font = New Font(Lbl得意先コード指定.Font, System.Drawing.FontStyle.Regular)
            Lbl得意先コード指定.ForeColor = Color.Black
            Lbl需要先コード指定.Font = New Font(Lbl需要先コード指定.Font, System.Drawing.FontStyle.Regular)
            Lbl需要先コード指定.ForeColor = Color.Black


            xxxlng得意先コード = 0
            xxxlng需要先コード = 0

            Bt印刷.Enabled = False

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　ヘッダコンボボックス選択時処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　sender -- オブジェクト識別クラス
    ' *　引数２　　　　　：　e      -- イベントデータクラス
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub HeaderCmb_SelectedValueChanged(sender As Object, e As EventArgs)

        Dim Tag As String
        Dim i As Integer

        Try

            Tag = CStr(CType(sender, ComboBox).Tag)

            Select Case Tag

                Case "ID1" '事業所コンボボックス

                    With DirectCast(Cmb事業所.SelectedItem, 事業所一覧)
                        My.Settings.事業所コード = .int事業所コード
                    End With

                    Cmb汎用.Items.Clear()
                    Cmb汎用.Text = ""
                    xxxstrProgram_ID = ""

                    InitializeDetail()

                    'サブプログラム一覧
                    If My.Settings.事業所コード <> 0 Then

                        If DTNP0000_PROC0001(gintSQLCODE, gstrSQLERRM) = False Then
                            MsgBox(gintSQLCODE & "-" & gstrSQLERRM & vbCr & MSG_COM902 & vbCr & MSG_COM901, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                            log.Error(Set_ErrMSG(gintSQLCODE, gstrSQLERRM))
                            Exit Sub
                        End If

                        For i = 0 To gintサブプログラムCnt - 1

                            Cmb汎用.Items.Add(New サブプログラム一覧(サブプログラムArray(i).strサブプログラム名, サブプログラムArray(i).strサブプログラムコード))

                        Next

                        '空白行追加
                        Cmb汎用.Items.Add(New サブプログラム一覧("", ""))

                    End If

                    Exit Select

                Case "ID2" 'サブプログラム

                    With DirectCast(Cmb汎用.SelectedItem, サブプログラム一覧)
                        xxxstrProgram_ID = .strサブプログラムコード
                    End With

                    InitializeDetail()

                    'サブプログラム情報
                    If IsNull(xxxstrProgram_ID) = False Then

                        If DTNP0000_PROC0002(xxxstrProgram_ID, gintSQLCODE, gstrSQLERRM) = False Then
                            MsgBox(gintSQLCODE & "-" & gstrSQLERRM & vbCr & MSG_COM902 & vbCr & MSG_COM901, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                            log.Error(Set_ErrMSG(gintSQLCODE, gstrSQLERRM))
                            Exit Sub
                        End If

                        gintRtn = DLTP0996S_PROC0001(xxxstrProgram_ID, 99, 0, gintSQLCODE, gstrSQLERRM)

                        If gudt帳票管理情報.lng帳票管理番号 > 0 Then Bt印刷.Enabled = True

                        Set検索条件()

                    End If

                    Exit Select

            End Select

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　Set検索条件
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　なし
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Set検索条件()

        Try

            Select Case gudtプログラムマスタ.int検索条件１ '得意先
                Case 0
                    Cmb得意先.Enabled = False
                    Lbl得意先.Font = New Font(Lbl得意先.Font, System.Drawing.FontStyle.Regular)
                    Lbl得意先.ForeColor = Color.Black
                Case 1
                    Cmb得意先.Enabled = True
                    Lbl得意先.Font = New Font(Lbl得意先.Font, System.Drawing.FontStyle.Bold)
                    Lbl得意先.ForeColor = Color.Red
                Case 2
                    Cmb得意先.Enabled = True
                    Lbl得意先.Font = New Font(Lbl得意先.Font, System.Drawing.FontStyle.Bold)
                    Lbl得意先.ForeColor = Color.Blue
            End Select

            Select Case gudtプログラムマスタ.int検索条件２ '需要先
                Case 0
                    Cmb需要先.Enabled = False
                    Lbl需要先.Font = New Font(Lbl需要先.Font, System.Drawing.FontStyle.Regular)
                    Lbl需要先.ForeColor = Color.Black
                Case 1
                    Cmb需要先.Enabled = True
                    Lbl需要先.Font = New Font(Lbl需要先.Font, System.Drawing.FontStyle.Bold)
                    Lbl需要先.ForeColor = Color.Red
                Case 2
                    Cmb需要先.Enabled = True
                    Lbl需要先.Font = New Font(Lbl需要先.Font, System.Drawing.FontStyle.Bold)
                    Lbl需要先.ForeColor = Color.Blue
            End Select

            Select Case gudtプログラムマスタ.int検索条件３ '対処日
                Case 0
                    txtDate.Enabled = False
                    Lbl対象日.Font = New Font(Lbl対象日.Font, System.Drawing.FontStyle.Regular)
                    Lbl対象日.ForeColor = Color.Black
                Case 1
                    txtDate.Enabled = True
                    Lbl対象日.Font = New Font(Lbl対象日.Font, System.Drawing.FontStyle.Bold)
                    Lbl対象日.ForeColor = Color.Red
                Case 2
                    txtDate.Enabled = True
                    Lbl対象日.Font = New Font(Lbl対象日.Font, System.Drawing.FontStyle.Bold)
                    Lbl対象日.ForeColor = Color.Blue
            End Select

            Select Case gudtプログラムマスタ.int検索条件４ '商品コード
                Case 0
                    txt商品コード.Enabled = False
                    txt商品コード.AlternateText.DisplayNull.Text = ""
                    Lbl商品.Font = New Font(Lbl商品.Font, System.Drawing.FontStyle.Regular)
                    Lbl商品.ForeColor = Color.Black
                Case 1
                    txt商品コード.Enabled = True
                    txt商品コード.AlternateText.DisplayNull.Text = gudtプログラムマスタ.str検索条件４ヒント
                    Lbl商品.Font = New Font(Lbl商品.Font, System.Drawing.FontStyle.Bold)
                    Lbl商品.ForeColor = Color.Red
                Case 2
                    txt商品コード.Enabled = True
                    txt商品コード.AlternateText.DisplayNull.Text = gudtプログラムマスタ.str検索条件４ヒント
                    Lbl商品.Font = New Font(Lbl商品.Font, System.Drawing.FontStyle.Bold)
                    Lbl商品.ForeColor = Color.Blue
            End Select

            Select Case gudtプログラムマスタ.int検索条件５ '相手先品番
                Case 0
                    txt相手先品番.Enabled = False
                    txt相手先品番.AlternateText.DisplayNull.Text = ""
                    Lbl相手先品番.Font = New Font(Lbl相手先品番.Font, System.Drawing.FontStyle.Regular)
                    Lbl相手先品番.ForeColor = Color.Black
                Case 1
                    txt相手先品番.Enabled = True
                    txt相手先品番.AlternateText.DisplayNull.Text = gudtプログラムマスタ.str検索条件５ヒント
                    Lbl相手先品番.Font = New Font(Lbl相手先品番.Font, System.Drawing.FontStyle.Bold)
                    Lbl相手先品番.ForeColor = Color.Red
                Case 2
                    txt相手先品番.Enabled = True
                    txt相手先品番.AlternateText.DisplayNull.Text = gudtプログラムマスタ.str検索条件５ヒント
                    Lbl相手先品番.Font = New Font(Lbl相手先品番.Font, System.Drawing.FontStyle.Bold)
                    Lbl相手先品番.ForeColor = Color.Blue
            End Select

            Select Case gudtプログラムマスタ.int検索条件６ '得意先コード
                Case 0
                    txt得意先.Enabled = False
                    txt得意先.AlternateText.DisplayNull.Text = ""
                    Lbl得意先コード指定.Font = New Font(Lbl得意先コード指定.Font, System.Drawing.FontStyle.Regular)
                    Lbl得意先コード指定.ForeColor = Color.Black
                Case 1
                    txt得意先.Enabled = True
                    txt得意先.AlternateText.DisplayNull.Text = gudtプログラムマスタ.str検索条件６ヒント
                    Lbl得意先コード指定.Font = New Font(Lbl得意先コード指定.Font, System.Drawing.FontStyle.Bold)
                    Lbl得意先コード指定.ForeColor = Color.Red
                Case 2
                    txt得意先.Enabled = True
                    txt得意先.AlternateText.DisplayNull.Text = gudtプログラムマスタ.str検索条件６ヒント
                    Lbl得意先コード指定.Font = New Font(Lbl得意先コード指定.Font, System.Drawing.FontStyle.Bold)
                    Lbl得意先コード指定.ForeColor = Color.Blue
            End Select

            'Select Case gudtプログラムマスタ.int検索条件６ '得意先コード
            '    Case 0
            '        txt得意先.Enabled = False
            '        txt得意先.AlternateText.DisplayNull.Text = ""
            '        Lbl得意先コード指定.Font = New Font(Lbl得意先コード指定.Font, System.Drawing.FontStyle.Regular)
            '        Lbl得意先コード指定.ForeColor = Color.Black
            '    Case 1
            '        txt得意先.Enabled = True
            '        txt得意先.AlternateText.DisplayNull.Text = gudtプログラムマスタ.str検索条件６ヒント
            '        Lbl得意先コード指定.Font = New Font(Lbl得意先コード指定.Font, System.Drawing.FontStyle.Bold)
            '        Lbl得意先コード指定.ForeColor = Color.Red
            '    Case 2
            '        txt得意先.Enabled = True
            '        txt得意先.AlternateText.DisplayNull.Text = gudtプログラムマスタ.str検索条件６ヒント
            '        Lbl得意先コード指定.Font = New Font(Lbl得意先コード指定.Font, System.Drawing.FontStyle.Bold)
            '        Lbl得意先コード指定.ForeColor = Color.Blue
            'End Select

            Select Case gudtプログラムマスタ.int検索条件７ '需要先コード
                Case 0
                    txt需要先.Enabled = False
                    txt需要先.AlternateText.DisplayNull.Text = ""
                    Lbl需要先コード指定.Font = New Font(Lbl需要先コード指定.Font, System.Drawing.FontStyle.Regular)
                    Lbl需要先コード指定.ForeColor = Color.Black
                Case 1
                    txt需要先.Enabled = True
                    txt需要先.AlternateText.DisplayNull.Text = gudtプログラムマスタ.str検索条件７ヒント
                    Lbl需要先コード指定.Font = New Font(Lbl需要先コード指定.Font, System.Drawing.FontStyle.Bold)
                    Lbl需要先コード指定.ForeColor = Color.Red
                Case 2
                    txt需要先.Enabled = True
                    txt需要先.AlternateText.DisplayNull.Text = gudtプログラムマスタ.str検索条件７ヒント
                    Lbl需要先コード指定.Font = New Font(Lbl需要先コード指定.Font, System.Drawing.FontStyle.Bold)
                    Lbl需要先コード指定.ForeColor = Color.Blue
            End Select


            If gudtプログラムマスタ.int検索条件１ > 0 Then

                If DTNP0000_PROC0003(xxxstrProgram_ID, 1, gintSQLCODE, gstrSQLERRM) = False Then
                    MsgBox(gintSQLCODE & "-" & gstrSQLERRM & vbCr & MSG_COM902 & vbCr & MSG_COM901, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                    log.Error(Set_ErrMSG(gintSQLCODE, gstrSQLERRM))
                    Exit Sub
                End If

                For i = 0 To gint得意先Cnt - 1

                    Cmb得意先.Items.Add(New 得意先一覧(得意先Array(i).str得意先名, 得意先Array(i).lng得意先コード))

                Next

                '空白行追加
                Cmb得意先.Items.Add(New 得意先一覧("", 0))

            End If


            If gudtプログラムマスタ.int検索条件２ > 0 Then

                If DTNP0000_PROC0003(xxxstrProgram_ID, 2, gintSQLCODE, gstrSQLERRM) = False Then
                    MsgBox(gintSQLCODE & "-" & gstrSQLERRM & vbCr & MSG_COM902 & vbCr & MSG_COM901, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                    log.Error(Set_ErrMSG(gintSQLCODE, gstrSQLERRM))
                    Exit Sub
                End If

                For i = 0 To gint需要先Cnt - 1

                    Cmb需要先.Items.Add(New 需要先一覧(需要先Array(i).str需要先名, 需要先Array(i).lng需要先コード))

                Next

                '空白行追加
                Cmb需要先.Items.Add(New 需要先一覧("", 0))

            End If

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　検索条件コンボボックス選択時処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　sender -- オブジェクト識別クラス
    ' *　引数２　　　　　：　e      -- イベントデータクラス
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Cmb_SelectedValueChanged(sender As Object, e As EventArgs)

        Dim Tag As String

        Try

            Tag = CStr(CType(sender, ComboBox).Tag)

            Select Case Tag

                Case "ID1" '得意先

                    With DirectCast(Cmb得意先.SelectedItem, 得意先一覧)
                        xxxlng得意先コード = .lng得意先コード
                    End With

                    Exit Select

                Case "ID2" '需要先

                    With DirectCast(Cmb需要先.SelectedItem, 需要先一覧)
                        xxxlng需要先コード = .lng需要先コード
                    End With

                    Exit Select

            End Select

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　クリアボタン押下処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　sender -- オブジェクト識別クラス
    ' *　引数２　　　　　：　e      -- イベントデータクラス
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Btクリア_Click(sender As Object, e As EventArgs) Handles Btクリア.Click

        Try

            gintRtn = MsgBox(MSG_COM001, CType(MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Question, MsgBoxStyle), My.Application.Info.Title)

            If gintRtn = vbYes Then

                Cmb汎用.Text = ""
                xxxstrProgram_ID = ""

                InitializeDetail()

                Cmb汎用.Focus()

            End If

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　検索ボタン押下処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　sender -- オブジェクト識別クラス
    ' *　引数２　　　　　：　e      -- イベントデータクラス
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Bt検索_Click(sender As Object, e As EventArgs) Handles Bt検索.Click

        Try

            'Oracle接続状態確認
            If (OraConnectState(gintRtn, gintSQLCODE, gstrSQLERRM)) = False Then
                MsgBox(gintSQLCODE & "-" & gstrSQLERRM & vbCr & MSG_COM004 & vbCr & MSG_COM903, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                Exit Sub
            End If

            If IsNull(Cmb事業所.Text.Trim) Then
                MsgBox(MSG_COM007, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                Cmb事業所.Focus()
                Exit Sub
            End If

            If IsNull(Cmb汎用.Text.Trim) Then
                MsgBox(MSG_COM012, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                Cmb汎用.Focus()
                Exit Sub
            End If

            '得意先
            If gudtプログラムマスタ.int検索条件１ = 1 AndAlso IsNull(Cmb得意先.Text.Trim) Then
                MsgBox(MSG_COM013, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                Cmb得意先.Focus()
                Exit Sub
            End If

            '需要先
            If gudtプログラムマスタ.int検索条件２ = 1 AndAlso IsNull(Cmb需要先.Text.Trim) Then
                MsgBox(MSG_COM014, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                Cmb需要先.Focus()
                Exit Sub
            End If

            '対象日
            Select Case gudtプログラムマスタ.int検索条件３

                Case 1 '必須

                    If IsNull(txtDate.Text.Trim) Then
                        MsgBox(MSG_COM015, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                        txtDate.Focus()
                        Exit Sub
                    End If
                    If Date_Check(txtDate.Text.Trim, 1) = False Then
                        txtDate.Text = ""
                        MsgBox(MSG_COM018, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                        txtDate.Focus()
                        Exit Sub
                    End If

                Case 2 '任意

                    If txtDate.Text.Trim.CompareTo(DUMMY_DATESTRING) = 1 Then '空白以外

                        If Date_Check(txtDate.Text.Trim, 1) = False Then
                            txtDate.Text = ""
                            MsgBox(MSG_COM018, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                            txtDate.Focus()
                            Exit Sub
                        End If

                    End If

            End Select

            '商品コード
            If gudtプログラムマスタ.int検索条件４ = 1 AndAlso IsNull(txt商品コード.Text.Trim) Then
                MsgBox(MSG_COM019, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                txt商品コード.Focus()
                Exit Sub
            End If

            '相手先品番
            If gudtプログラムマスタ.int検索条件５ = 1 AndAlso IsNull(txt相手先品番.Text.Trim) Then
                MsgBox(MSG_COM020, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                txt相手先品番.Focus()
                Exit Sub
            End If

            '得意先コード
            If gudtプログラムマスタ.int検索条件６ = 1 AndAlso IsNull(txt得意先.Text.Trim) Then
                MsgBox(MSG_COM022, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                txt得意先.Focus()
                Exit Sub
            End If

            '需要先コード
            If gudtプログラムマスタ.int検索条件７ = 1 AndAlso IsNull(txt需要先.Text.Trim) Then
                MsgBox(MSG_COM023, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                txt需要先.Focus()
                Exit Sub
            End If


            TSSRowCount.Text = MSG_COM021 & "0"

            'gblRtn = 実行処理(My.Settings.事業所コード, xxxstrProgram_ID)
            gblRtn = 実行処理(xxxstrProgram_ID)

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　フォームサイズ変更処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　sender・・オブジェクト識別クラス
    ' *　引数２　　　　　：　e・・イベントデータクラス
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Fm_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged

        Try

            'サイドバーコントロール位置調整
            Container検索.Height = Dgv.Height - 6
            Bt検索.Location = New Point(12, Dgv.Height - 70)
            Btクリア.Location = New Point(128, Dgv.Height - 70)
            Bt印刷.Location = New Point(12, Dgv.Height - 37)
            Lbl注意１.Location = New Point(16, Dgv.Height - 100)
            Lbl注意２.Location = New Point(102, Dgv.Height - 100)

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　実行処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　int事業所コード
    ' *　引数２　　　　　：　strProgram_ID
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    'Private Function 実行処理(ByVal int事業所コード As Integer, ByVal strProgram_ID As String) As Boolean
    Private Function 実行処理(ByVal strProgram_ID As String) As Boolean
        Try

            実行処理 = False

            Select Case strProgram_ID

                Case "DTN0001", "DTN0002" '出荷検品未完了(ロット管理のみ)、出荷検品未完了(全商品)

                    Cursor = Cursors.WaitCursor

                    gblRtn = DTNP0403_PROC001(gudtプログラムマスタ.str処理関数, gudtプログラムマスタ.intサブプログラム_ID, xxxlng得意先コード, xxxlng需要先コード, Dgv, gintResultCnt)

                    'セッション情報削除
                    gintRtn = DLTP0998S_PROC0013("HARKP201", "出荷検品未完了", gudtプログラムマスタ.intサブプログラム_ID)

                    Cursor = Cursors.Default

                    If gintResultCnt = 0 Then

                        MsgBox(MSG_COM002, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                        Cmb汎用.Focus()
                        Exit Function

                    End If

                    TSSRowCount.Text = MSG_COM021 & gintResultCnt

                    Exit Select


                Case "DTN0003", "DTN0004" '【天神会】SPDシステム受注エラー情報、【天神会】Oliver取込処理エラー情報

                    Cursor = Cursors.WaitCursor

                    gblRtn = DTNP0403_PROC003(gudtプログラムマスタ.str処理関数, txtDate.Text.Trim, Dgv, gintResultCnt)

                    'セッション情報削除
                    gintRtn = DLTP0998S_PROC0013("HARKP301", "Oliver取込処理エラー", 1)

                    Cursor = Cursors.Default

                    If gintResultCnt = 0 Then

                        MsgBox(MSG_COM002, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                        Cmb汎用.Focus()
                        Exit Function

                    End If

                    TSSRowCount.Text = MSG_COM021 & gintResultCnt

                    Exit Select

                Case "DTN0005" '【天神会】長期貸出番号情報

                    Cursor = Cursors.WaitCursor

                    gblRtn = DTNP0403_PROC004(gudtプログラムマスタ.str処理関数, xxxlng得意先コード, xxxlng需要先コード, txt商品コード.Text.Trim, txt相手先品番.Text.Trim, Dgv, gintResultCnt)

                    Cursor = Cursors.Default

                    If gintResultCnt = 0 Then

                        MsgBox(MSG_COM002, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                        Cmb汎用.Focus()
                        Exit Function

                    End If

                    TSSRowCount.Text = MSG_COM021 & gintResultCnt

                    Exit Select

                Case "DTN0006" '【共通】 PHsmos受注情報

                    Cursor = Cursors.WaitCursor


                    gblRtn = DTNP0403_PROC006(gudtプログラムマスタ.str処理関数, My.Settings.事業所コード, xxxlng得意先コード, txtDate.Text.Trim, Dgv, gintResultCnt)

                    Cursor = Cursors.Default

                    If gintResultCnt = 0 Then

                        MsgBox(MSG_COM002, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                        Cmb汎用.Focus()
                        Exit Function

                    End If

                    TSSRowCount.Text = MSG_COM021 & gintResultCnt

                    Exit Select

                Case "DTN0007" '【共通】 出荷検品未完了(全商品)

                    Cursor = Cursors.WaitCursor

                    gblRtn = DTNP0403_PROC001(gudtプログラムマスタ.str処理関数, gudtプログラムマスタ.intサブプログラム_ID, CLng(NvlString(txt得意先.Text.Trim, "0")), CLng(NvlString(txt需要先.Text.Trim, "0")), Dgv, gintResultCnt)

                    'セッション情報削除
                    gintRtn = DLTP0998S_PROC0013("HARKP201", "出荷検品未完了", gudtプログラムマスタ.intサブプログラム_ID)

                    Cursor = Cursors.Default

                    If gintResultCnt = 0 Then

                        MsgBox(MSG_COM002, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                        Cmb汎用.Focus()
                        Exit Function

                    End If

                    TSSRowCount.Text = MSG_COM021 & gintResultCnt

                    Exit Select

                Case "DTN0008" '需要先別売上金額

                    Cursor = Cursors.WaitCursor

                    gblRtn = DTNP0403_PROC007(gudtプログラムマスタ.str処理関数, My.Settings.事業所コード, CLng(NvlString(txt得意先.Text.Trim, "0")), Dgv, gintResultCnt)

                    'セッション情報削除
                    gintRtn = DLTP0998S_PROC0013("DTN0008", "需要先別売上情報", 1)

                    Cursor = Cursors.Default

                    If gintResultCnt = 0 Then

                        MsgBox(MSG_COM002, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                        Cmb汎用.Focus()
                        Exit Function

                    End If

                    TSSRowCount.Text = MSG_COM021 & gintResultCnt

                    Exit Select

                Case "DTN0009" '【天神会】有効期限切迫情報

                    Cursor = Cursors.WaitCursor


                    gblRtn = DTNP0403_PROC008(gudtプログラムマスタ.str処理関数, xxxlng得意先コード, txtDate.Text.Trim, Dgv, gintResultCnt)

                    Cursor = Cursors.Default

                    If gintResultCnt = 0 Then

                        MsgBox(MSG_COM002, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                        Cmb汎用.Focus()
                        Exit Function

                    End If

                    TSSRowCount.Text = MSG_COM021 & gintResultCnt

                    Exit Select

            End Select

            実行処理 = True

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *  モジュール機能　： キーダウン処理
    ' *
    ' *  注意、制限事項  ：なし
    ' *  引数　　　　　　：sender・・オブジェクト識別クラス
    ' *      　　　　　　：e・・イベントデータクラス
    ' *　戻値　　　　　　：なし
    ' *-----------------------------------------------------------------------------/ 
    Private Sub Txt_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)

        Try

            '[Enter]キー判別
            If e.KeyCode = Keys.Enter Then

                'TabIndexの次のコントロールへフォーカス移動
                Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
                e.Handled = True

            End If

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　印刷ボタン押下処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　sender -- オブジェクト識別クラス
    ' *　引数２　　　　　：　e      -- イベントデータクラス
    ' *　戻値　　　　　　：　なし
    ' *-----------------------------------------------------------------------------/
    Private Sub Bt印刷_Click(sender As Object, e As EventArgs) Handles Bt印刷.Click

        Dim Doc As Document = Nothing
        Dim i As Integer
        Dim blPrintFlg As Boolean
        Dim intPrintChk As Integer

        Try

            If Viewer Is Nothing Then Viewer = New HARK990()

            If Dgv.RowCount < 1 Then

                MsgBox(MSG_COM002, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                Bt印刷.Focus()
                Exit Sub

            End If

            Select Case xxxstrProgram_ID

                '出荷検品未完了
                Case "DTN0001", "DTN0002", "DTN0007"

                    If DTN0001_CreateDocument(Dgv, VBReport, Doc) = False Then
                        MsgBox(MSG_COM803, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)
                        Bt印刷.Focus()
                        Exit Sub
                    End If

                    Exit Select

                'PHsmos受注情報
                Case "DTN0006"

                    If DTN0006_CreateDocument(Dgv, VBReport, Doc) = False Then
                        MsgBox(MSG_COM803, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)
                        Bt印刷.Focus()
                        Exit Sub
                    End If

                    Exit Select

                '有効期限切迫情報
                Case "DTN0009"

                    blPrintFlg = False

                    For i = 0 To Dgv.RowCount - 1

                        '印刷チェックボックス
                        If IsDBNull(Dgv.Rows(i).Cells(24).Value) = True Then
                            intPrintChk = 0
                        Else
                            intPrintChk = CInt(Dgv.Rows(i).Cells(24).Value)
                        End If

                        '未印刷＆印刷チェックボックスON
                        If CInt(Dgv.Rows(i).Cells(21).Value) = 0 AndAlso intPrintChk <> 0 Then
                            blPrintFlg = True
                        End If

                    Next

                    If blPrintFlg = False Then
                        MsgBox(MSG_COM002, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                        Bt印刷.Focus()
                        Exit Sub
                    End If

                    If DTN0009_CreateDocument(Dgv, VBReport, Doc) = False Then
                        MsgBox(MSG_COM803, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)
                        Bt印刷.Focus()
                        Exit Sub
                    End If

            End Select

            If Doc IsNot Nothing Then

                Viewer.ReportDocument = Doc
                Viewer.ShowDialog()

            Else

                MsgBox(MSG_COM002, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                Bt印刷.Focus()
                Exit Sub

            End If

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        Finally

            If Not Viewer Is Nothing Then Viewer.Dispose()

        End Try

    End Sub

End Class
