'/*-----------------------------------------------------------------------------
' * COPYRIGHT(C) ITI CORPORATION 2020
' * ITI CONFIDENTIAL AND PROPRIETARY
' *
' * All rights reserved by ITI Corporation.
' *-----------------------------------------------------------------------------/
Option Compare Binary
Option Explicit On
Option Strict On

Imports Oracle.DataAccess.Client
'Imports Oracle.DataAccess.Types
Imports HARK010.HARK_Common
Imports HARK010.HARK_Sub

Public Class HARK_DBCommon

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Shared Oracomm As New OracleConnection
    Public Shared Oracmd As New OracleCommand
    Public Shared OraDr As OracleDataReader
    Public Shared OraDar As OracleDataAdapter
    Public Shared OraTran As OracleTransaction

    '/*-----------------------------------------------------------------------------
    ' *　モジュール名　　：　OraConnect
    ' *　クラス名　　　　：　HASS_DBCommon
    ' *　モジュール機能　：　Oracle接続処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　なし
    ' *　戻値　　　　　　：　True・・成功、false・・失敗
    ' *
    ' *　作成者　　　　　：　k.takada
    ' *　作成日　　　　　：　2006.3.5
    ' *　修正履歴　　　　：　
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function OraConnect() As Boolean

        Dim StrParam As String

        Try

            OraConnect = False

            StrParam = "User id=" & My.Settings.DBユーザ & ";" & "Password=" & My.Settings.DBパスワード & ";" & "Data Source=" & My.Settings.DB接続文字列

            Oracomm.ConnectionString = StrParam

            Oracomm.Open()

            OraConnect = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール名　　：　OraDisConnect
    ' *　クラス名　　　　：　HASS_DBCommon
    ' *　モジュール機能　：　Oracle切断処理
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　なし
    ' *　戻値　　　　　　：　True・・成功、false・・失敗
    ' *
    ' *　作成者　　　　　：　k.takada
    ' *　作成日　　　　　：　2006.3.5
    ' *　修正履歴　　　　：　
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function OraDisConnect() As Boolean

        Try

            OraDisConnect = False

            Oracomm.Close()

            OraDisConnect = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()
            Oracomm.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール名　　：　OraConnectState
    ' *　クラス名　　　　：　HASS_DBCommon
    ' *　モジュール機能　：　Oracle接続状態確認
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　OraSessionID   -- セッションID（戻値）
    ' *　引数２　　　　　：　SQLCODE        -- Oracleエラーコード（戻値）
    ' *　引数３　　　　　：　SQLERRM        -- Oracleエラーメッセージ（戻値）
    ' *　戻値　　　　　　：　True・・接続中、false・・切断
    ' *-----------------------------------------------------------------------------/
    Public Shared Function OraConnectState(ByRef PO_intOraSessionID As Integer,
                                           ByRef PO_intSQLCODE As Integer,
                                           ByRef PO_strSQLERRM As String) As Boolean

        Try

            Dim PO_01 As OracleParameter
            Dim PO_02 As OracleParameter
            Dim PO_03 As OracleParameter
            Dim PO_04 As OracleParameter

            OraConnectState = False
            PO_intOraSessionID = 0

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0998S.PROC0001"
            Oracmd.CommandType = CommandType.StoredProcedure

            'Outputパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.Varchar2, 255, DBNull.Value, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Int32, ParameterDirection.Output)
            PO_04 = Oracmd.Parameters.Add("PO_04", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            'ストアドプロシージャCall
            Oracmd.ExecuteNonQuery()

            PO_intSQLCODE = CInt(PO_03.Value.ToString)
            PO_strSQLERRM = PO_04.Value.ToString

            'リターンコードでの処理振り分け
            If PO_intSQLCODE = 0 Then

                PO_intOraSessionID = CInt(PO_02.Value.ToString)
                OraConnectState = True

            Else

                log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

            End If

            Exit Function

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))

            Return OraConnectState

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　事業所一覧取得
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　Program_ID   -- プログラム_ID
    ' *　引数２　　　　　：　SQLCODE      -- Oracleエラーコード（戻値）
    ' *　引数３　　　　　：　SQLERRM      -- Oracleエラーメッセージ（戻値）
    ' *　戻値　　　　　　：　True -- 正常取得 False -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DLTP0901_PROC0001(ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Dim i As Integer

        Try
            DLTP0901_PROC0001 = False

            gint事業所Cnt = 0

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0901.PROC0001"
            Oracmd.CommandType = CommandType.StoredProcedure

            'Outputパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            'ストアドプロシージャCall
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            事業所Array = Nothing
            gint事業所Cnt = 0

            'リターンコードでの処理振り分け
            If PO_intSQLCODE = 0 Then

                i = 0

                While OraDr.Read

                    'メモリ再取得
                    ReDim Preserve 事業所Array(i)

                    'グローバル変数にセット
                    事業所Array(i).int事業所コード = OraDr.GetInt32(0)
                    事業所Array(i).str事業所名 = OraDr.GetString(1)
                    i += 1

                End While

                gint事業所Cnt = i

            Else

                log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                Exit Function

            End If

            DLTP0901_PROC0001 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　部門名取得
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　Program_ID   -- プログラム_ID
    ' *　引数２　　　　　：　SQLCODE      -- Oracleエラーコード（戻値）
    ' *　引数３　　　　　：　SQLERRM      -- Oracleエラーメッセージ（戻値）
    ' *　戻値　　　　　　：　True -- 正常取得 False -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DLTP0900_PROC0002(ByVal PI_strProgram_ID As String,
                                             ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim PI_01 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Try
            DLTP0900_PROC0002 = False

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0900.PROC0002"
            Oracmd.CommandType = CommandType.StoredProcedure

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)

            'インプット値設定
            PI_01.Value = My.Settings.事業所コード

            'Outputパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.Varchar2, 60, DBNull.Value, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            'ストアドプロシージャCall
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            gstr部門名 = Nothing

            'リターンコードでの処理振り分け
            If PO_intSQLCODE = 0 Then

                gstr部門名 = PO_01.Value.ToString

            Else

                log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                Exit Function

            End If

            DLTP0900_PROC0002 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　サブプログラム一覧取得
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　SQLCODE      -- Oracleエラーコード（戻値）
    ' *　引数２　　　　　：　SQLERRM      -- Oracleエラーメッセージ（戻値）
    ' *　戻値　　　　　　：　True -- 正常取得 False -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0000_PROC0001(ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim PI_01 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Dim i As Integer

        Try
            DTNP0000_PROC0001 = False

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DTNP0000.PROC0001"
            Oracmd.CommandType = CommandType.StoredProcedure

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)

            'インプット値設定
            PI_01.Value = My.Settings.事業所コード

            'Outputパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            'ストアドプロシージャCall
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            サブプログラムArray = Nothing
            gintサブプログラムCnt = 0

            'リターンコードでの処理振り分け
            If PO_intSQLCODE = 0 Then

                i = 0

                While OraDr.Read

                    'メモリ再取得
                    ReDim Preserve サブプログラムArray(i)

                    'グローバル変数にセット
                    サブプログラムArray(i).strサブプログラムコード = OraDr.GetString(0)
                    サブプログラムArray(i).strサブプログラム名 = OraDr.GetString(1)
                    i += 1

                End While

                gintサブプログラムCnt = i

            Else

                log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                Exit Function

            End If

            DTNP0000_PROC0001 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：  プログラム情報取得
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　Program_ID   -- プログラム_ID
    ' *　引数２　　　　　：　SQLCODE      -- Oracleエラーコード（戻値）
    ' *　引数３　　　　　：　SQLERRM      -- Oracleエラーメッセージ（戻値）
    ' *　戻値　　　　　　：　True -- 正常取得 False -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0000_PROC0002(ByVal PI_strProgram_ID As String,
                                             ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Try
            DTNP0000_PROC0002 = False

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DTNP0000.PROC0002"
            Oracmd.CommandType = CommandType.StoredProcedure

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Varchar2, ParameterDirection.Input)

            'インプット値設定
            PI_01.Value = My.Settings.事業所コード
            PI_02.Value = PI_strProgram_ID

            'Outputパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            'ストアドプロシージャCall
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            gudtプログラムマスタ.IsClear()

            'リターンコードでの処理振り分け
            If PO_intSQLCODE = 0 Then

                OraDr.Read()

                If OraDr.IsDBNull(0) = False Then gudtプログラムマスタ.str処理関数 = OraDr.GetString(0)
                If OraDr.IsDBNull(1) = False Then gudtプログラムマスタ.str出力ヘッダ = OraDr.GetString(1)
                If OraDr.IsDBNull(2) = False Then gudtプログラムマスタ.str出力区切文字 = OraDr.GetString(2)
                If OraDr.IsDBNull(3) = False Then gudtプログラムマスタ.int検索条件１ = OraDr.GetInt32(3)
                If OraDr.IsDBNull(4) = False Then gudtプログラムマスタ.int検索条件２ = OraDr.GetInt32(4)
                If OraDr.IsDBNull(5) = False Then gudtプログラムマスタ.int検索条件３ = OraDr.GetInt32(5)
                If OraDr.IsDBNull(6) = False Then gudtプログラムマスタ.int検索条件４ = OraDr.GetInt32(6)
                If OraDr.IsDBNull(7) = False Then gudtプログラムマスタ.int検索条件５ = OraDr.GetInt32(7)
                If OraDr.IsDBNull(8) = False Then gudtプログラムマスタ.int検索条件６ = OraDr.GetInt32(8)
                If OraDr.IsDBNull(9) = False Then gudtプログラムマスタ.int検索条件７ = OraDr.GetInt32(9)
                If OraDr.IsDBNull(10) = False Then gudtプログラムマスタ.int検索条件８ = OraDr.GetInt32(10)
                If OraDr.IsDBNull(11) = False Then gudtプログラムマスタ.int検索条件９ = OraDr.GetInt32(11)
                If OraDr.IsDBNull(12) = False Then gudtプログラムマスタ.int検索条件１０ = OraDr.GetInt32(12)
                If OraDr.IsDBNull(13) = False Then gudtプログラムマスタ.str検索条件１ヒント = OraDr.GetString(13)
                If OraDr.IsDBNull(14) = False Then gudtプログラムマスタ.str検索条件２ヒント = OraDr.GetString(14)
                If OraDr.IsDBNull(15) = False Then gudtプログラムマスタ.str検索条件３ヒント = OraDr.GetString(15)
                If OraDr.IsDBNull(16) = False Then gudtプログラムマスタ.str検索条件４ヒント = OraDr.GetString(16)
                If OraDr.IsDBNull(17) = False Then gudtプログラムマスタ.str検索条件５ヒント = OraDr.GetString(17)
                If OraDr.IsDBNull(18) = False Then gudtプログラムマスタ.str検索条件６ヒント = OraDr.GetString(18)
                If OraDr.IsDBNull(19) = False Then gudtプログラムマスタ.str検索条件７ヒント = OraDr.GetString(19)
                If OraDr.IsDBNull(20) = False Then gudtプログラムマスタ.str検索条件８ヒント = OraDr.GetString(20)
                If OraDr.IsDBNull(21) = False Then gudtプログラムマスタ.str検索条件９ヒント = OraDr.GetString(21)
                If OraDr.IsDBNull(22) = False Then gudtプログラムマスタ.str検索条件１０ヒント = OraDr.GetString(22)
                If OraDr.IsDBNull(23) = False Then gudtプログラムマスタ.intサブプログラム_ID = OraDr.GetInt32(23)

            Else

                log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                Exit Function

            End If

            DTNP0000_PROC0002 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　名称辞書一覧取得
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　Program_ID   -- プログラム_ID
    ' *　引数２　　　　　：　区分         -- 条件識別区分
    ' *　引数３　　　　　：　SQLCODE      -- Oracleエラーコード（戻値）
    ' *　引数４　　　　　：　SQLERRM      -- Oracleエラーメッセージ（戻値）
    ' *　戻値　　　　　　：　True -- 正常取得 False -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0000_PROC0003(ByVal PI_strProgram_ID As String,
                                             ByVal PI_区分 As Integer,
                                             ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Dim i As Integer

        Try
            DTNP0000_PROC0003 = False

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DTNP0000.PROC0003"
            Oracmd.CommandType = CommandType.StoredProcedure

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Varchar2, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Int32, ParameterDirection.Input)

            'インプット値設定
            PI_01.Value = My.Settings.事業所コード
            PI_02.Value = PI_strProgram_ID
            PI_03.Value = PI_区分

            'Outputパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            'ストアドプロシージャCall
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            Select Case PI_区分

                Case 1

                    得意先Array = Nothing
                    gint得意先Cnt = 0

                    'リターンコードでの処理振り分け
                    If PO_intSQLCODE = 0 Then

                        i = 0

                        While OraDr.Read

                            'メモリ再取得
                            ReDim Preserve 得意先Array(i)

                            'グローバル変数にセット
                            得意先Array(i).lng得意先コード = OraDr.GetInt64(0)
                            得意先Array(i).str得意先名 = OraDr.GetString(1)
                            i += 1

                        End While

                        gint得意先Cnt = i

                    Else

                        log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                        Exit Function

                    End If

                Case 2

                    需要先Array = Nothing
                    gint需要先Cnt = 0

                    'リターンコードでの処理振り分け
                    If PO_intSQLCODE = 0 Then

                        i = 0

                        While OraDr.Read

                            'メモリ再取得
                            ReDim Preserve 需要先Array(i)

                            'グローバル変数にセット
                            需要先Array(i).lng需要先コード = OraDr.GetInt64(0)
                            需要先Array(i).str需要先名 = OraDr.GetString(1)
                            i += 1

                        End While

                        gint需要先Cnt = i

                    Else

                        log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                        Exit Function

                    End If

            End Select

            DTNP0000_PROC0003 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　出荷検品未完了データ検索
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　strProgram_ID     -- プログラム_ID
    ' *  引数２　　　　　：　サブプログラム_ID -- サブプログラム_ID
    ' *  引数３　　　　　：　得意先コード      -- 得意先コード
    ' *  引数４　　　　　：　需要先コード      -- 需要先コード
    ' *　引数５　　　　　：　Dgv               -- DataGridView（戻値）
    ' *　引数６　　　　　：　ROWCount          -- 件数（戻値）
    ' *　戻値　　　　　　：　0 -- 正常取得 2 -- レコード無 9 -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC001(ByVal PI_strProgram_ID As String,
                                            ByVal PI_intサブプログラム_ID As Integer,
                                            ByVal PI_lng得意先コード As Long,
                                            ByVal PI_lng需要先コード As Long,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PI_04 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC001 = False

            OraDs.Clear()

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            OraTran = Oracomm.BeginTransaction

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int64, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Int64, ParameterDirection.Input)
            PI_04 = Oracmd.Parameters.Add("PI_04", OracleDbType.Int32, ParameterDirection.Input)

            'インプット値設定
            PI_01.Value = PI_intサブプログラム_ID
            PI_02.Value = PI_lng得意先コード
            If PI_lng需要先コード = 0 Then
                PI_03.Value = vbNullString
            Else
                PI_03.Value = PI_lng需要先コード
            End If
            PI_04.Value = My.Settings.事業所コード

            'アウトプットパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            'ストアドプロシージャcall
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            OraTran.Commit()

            DTNP0403_PROC001 = True

        Catch Oraex As OracleException

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            If Not IsNothing(OraTran) Then
                OraTran.Dispose()
            End If

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　Oliverエラーデータ検索
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　strProgram_ID     -- プログラム_ID
    ' *  引数２　　　　　：　対象日            -- 検索対象日
    ' *　引数３　　　　　：　Dgv               -- DataGridView（戻値）
    ' *　引数４　　　　　：　ROWCount          -- 件数（戻値）
    ' *　戻値　　　　　　：　0 -- 正常取得 2 -- レコード無 9 -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC003(ByVal PI_strProgram_ID As String,
                                            ByVal PI_対象日 As String,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC003 = False

            OraDs.Clear()

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            OraTran = Oracomm.BeginTransaction

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)

            'インプット値設定
            PI_01.Value = PI_対象日.Trim

            'アウトプットパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            'ストアドプロシージャcall
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            OraTran.Commit()

            DTNP0403_PROC003 = True

        Catch Oraex As OracleException

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            If Not IsNothing(OraTran) Then
                OraTran.Dispose()
            End If

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　長期貸出番号情報検索
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　strProgram_ID     -- プログラム_ID
    ' *  引数２　　　　　：　得意先コード      -- 得意先コード
    ' *  引数３　　　　　：　需要先コード      -- 需要先コード
    ' *  引数４　　　　　：　商品コード        -- 商品コード
    ' *  引数５　　　　　：　相手先品番        -- 相手先品番
    ' *　引数６　　　　　：　Dgv               -- DataGridView（戻値）
    ' *　引数７　　　　　：　ROWCount          -- 件数（戻値）
    ' *　戻値　　　　　　：　0 -- 正常取得 2 -- レコード無 9 -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC004(ByVal PI_strProgram_ID As String,
                                            ByVal PI_lng得意先コード As Long,
                                            ByVal PI_lng需要先コード As Long,
                                            ByVal PI_str商品コード As String,
                                            ByVal PI_str相手先品番 As String,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PI_04 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC004 = False

            OraDs.Clear()

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int64, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int64, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Varchar2, 60, DBNull.Value, ParameterDirection.Input)
            PI_04 = Oracmd.Parameters.Add("PI_04", OracleDbType.Varchar2, 60, DBNull.Value, ParameterDirection.Input)

            'インプット値設定
            PI_01.Value = PI_lng得意先コード
            PI_02.Value = PI_lng需要先コード
            If IsNull(PI_str商品コード) Then
                PI_03.Value = vbNullString
            Else
                PI_03.Value = PI_str商品コード
            End If
            If IsNull(PI_str相手先品番) Then
                PI_04.Value = vbNullString
            Else
                PI_04.Value = PI_str相手先品番
            End If

            'アウトプットパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            'ストアドプロシージャcall
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            DTNP0403_PROC004 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　PHsmos受注情報検索
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　strProgram_ID     -- プログラム_ID
    ' *  引数２　　　　　：　部署コード        -- 部署コード
    ' *  引数２　　　　　：　得意先コード      -- 得意先コード
    ' *  引数３　　　　　：　対象日            -- 対象日
    ' *　引数４　　　　　：　Dgv               -- DataGridView（戻値）
    ' *　引数５　　　　　：　ROWCount          -- 件数（戻値）
    ' *　戻値　　　　　　：　0 -- 正常取得 2 -- レコード無 9 -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC006(ByVal PI_strProgram_ID As String,
                                            ByVal PI_int部署コード As Integer,
                                            ByVal PI_lng得意先コード As Long,
                                            ByVal PI_対象日 As String,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC006 = False

            OraDs.Clear()

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int64, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)

            'インプット値設定
            PI_01.Value = PI_int部署コード
            PI_02.Value = PI_lng得意先コード
            If PI_対象日.Trim.CompareTo(DUMMY_DATESTRING) = 0 Then
                PI_03.Value = vbNullString
            Else
                PI_03.Value = PI_対象日.Trim
            End If

            'アウトプットパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            'ストアドプロシージャcall
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            DTNP0403_PROC006 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　セッション情報削除
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　strProgram_ID   -- プログラム_ID
    ' *　引数２　　　　　：　strV区分        -- SUBセッション情報.V区分
    ' *　引数３　　　　　：　intN区分        -- SUBセッション情報.N区分
    ' *　戻値　　　　　　：　0 -- 正常終了 9 -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DLTP0998S_PROC0013(ByVal strProgram_ID As String,
                                              ByVal strV区分 As String,
                                              ByVal intN区分 As Integer) As Integer


        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PI_04 As OracleParameter
        Dim PI_05 As OracleParameter
        Dim PI_06 As OracleParameter
        Dim PI_07 As OracleParameter
        Dim PI_08 As OracleParameter
        Dim PI_09 As OracleParameter
        Dim PI_10 As OracleParameter
        Dim PI_11 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_99 As OracleParameter

        Try

            DLTP0998S_PROC0013 = 9

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0998S.PROC0013"
            Oracmd.CommandType = CommandType.StoredProcedure

            'OraTran = Oracomm.BeginTransaction

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)
            PI_04 = Oracmd.Parameters.Add("PI_04", OracleDbType.Int32, ParameterDirection.Input)
            PI_05 = Oracmd.Parameters.Add("PI_05", OracleDbType.Int32, ParameterDirection.Input)
            PI_06 = Oracmd.Parameters.Add("PI_06", OracleDbType.Int32, ParameterDirection.Input)
            PI_07 = Oracmd.Parameters.Add("PI_07", OracleDbType.Int32, ParameterDirection.Input)
            PI_08 = Oracmd.Parameters.Add("PI_08", OracleDbType.Int32, ParameterDirection.Input)
            PI_09 = Oracmd.Parameters.Add("PI_09", OracleDbType.Int32, ParameterDirection.Input)
            PI_10 = Oracmd.Parameters.Add("PI_10", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)
            PI_11 = Oracmd.Parameters.Add("PI_11", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)

            'Outputパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.Varchar2, 255, DBNull.Value, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_99 = Oracmd.Parameters.Add("PO_99", OracleDbType.Int32, ParameterDirection.Output)

            'インプット値設定
            PI_01.Value = "DELETE"
            PI_02.Value = strProgram_ID
            PI_03.Value = strV区分
            PI_04.Value = intN区分
            PI_05.Value = vbNullString
            PI_06.Value = vbNullString
            PI_07.Value = vbNullString
            PI_08.Value = vbNullString
            PI_09.Value = vbNullString
            PI_10.Value = vbNullString
            PI_11.Value = vbNullString

            'ストアドプロシージャcall
            Oracmd.ExecuteNonQuery()

            'リターンコードでの処理振り分け
            If CInt(PO_99.Value.ToString) = 0 Then

                DLTP0998S_PROC0013 = 0

                Exit Function

            End If

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　需要先別売上情報出力
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　strProgram_ID     -- プログラム_ID
    ' *  引数２　　　　　：　部署コード        -- 部署コード
    ' *  引数３　　　　　：　得意先コード      -- 得意先コード
    ' *　引数４　　　　　：　Dgv               -- DataGridView（戻値）
    ' *　引数５　　　　　：　ROWCount          -- 件数（戻値）
    ' *　戻値　　　　　　：　0 -- 正常取得 2 -- レコード無 9 -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC007(ByVal PI_strProgram_ID As String,
                                            ByVal PI_int部署コード As Integer,
                                            ByVal PI_lng得意先コード As Long,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC007 = False

            OraDs.Clear()

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int64, ParameterDirection.Input)

            'インプット値設定
            PI_01.Value = PI_int部署コード
            PI_02.Value = PI_lng得意先コード

            'アウトプットパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            'ストアドプロシージャcall
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            DTNP0403_PROC007 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　有効期限切迫情報検索
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　strProgram_ID     -- プログラム_ID
    ' *  引数２　　　　　：　得意先コード      -- 得意先コード
    ' *  引数３　　　　　：　対象日            -- 対象日
    ' *　引数４　　　　　：　Dgv               -- DataGridView（戻値）
    ' *　引数５　　　　　：　ROWCount          -- 件数（戻値）
    ' *　戻値　　　　　　：　0 -- 正常取得 2 -- レコード無 9 -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC008(ByVal PI_strProgram_ID As String,
                                            ByVal PI_lng得意先コード As Long,
                                            ByVal PI_対象日 As String,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC008 = False

            OraDs.Clear()

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int64, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)

            'インプット値設定
            PI_01.Value = PI_lng得意先コード
            PI_02.Value = PI_対象日.Trim

            'アウトプットパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            'ストアドプロシージャcall
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            '印刷チェックボックス追加
            OraDs.Tables("TMP").Columns.Add("印刷", Type.GetType("System.Boolean")).DefaultValue = False

            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            DTNP0403_PROC008 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：  帳票管理情報取得
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　Program_ID         -- プログラム_ID
    ' *　引数２　　　　　：　SPDシステムコード  -- SPDシステムコード
    ' *　引数３　　　　　：　サブプログラム_ID  -- サブプログラム_ID
    ' *　引数４　　　　　：　SQLCODE            -- Oracleエラーコード（戻値）
    ' *　引数５　　　　　：　SQLERRM            -- Oracleエラーメッセージ（戻値）
    ' *　戻値　　　　　　：　0 -- 正常取得 1 -- データ無し 9 -- エラー
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DLTP0996S_PROC0001(ByVal PI_strProgram_ID As String,
                                             ByVal PI_intSPDシステムコード As Integer,
                                             ByVal PI_intサブプログラム_ID As Integer,
                                             ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Integer

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PI_04 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Try
            DLTP0996S_PROC0001 = 9

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0996S.PROC0001"
            Oracmd.CommandType = CommandType.StoredProcedure

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Varchar2, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int32, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Int32, ParameterDirection.Input)
            PI_04 = Oracmd.Parameters.Add("PI_04", OracleDbType.Int32, ParameterDirection.Input)


            'インプット値設定
            PI_01.Value = PI_strProgram_ID
            PI_02.Value = PI_intSPDシステムコード
            PI_03.Value = PI_intサブプログラム_ID
            PI_04.Value = My.Settings.事業所コード

            'Outputパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            'ストアドプロシージャCall
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            gudt帳票管理情報.IsClear()

            'リターンコードでの処理振り分け
            Select Case PO_intSQLCODE

                Case 0

                    OraDr.Read()

                    If OraDr.IsDBNull(0) = False Then gudt帳票管理情報.lng帳票管理番号 = OraDr.GetInt64(0)
                    If OraDr.IsDBNull(1) = False Then gudt帳票管理情報.str帳票名 = OraDr.GetString(1)
                    If OraDr.IsDBNull(2) = False Then gudt帳票管理情報.strテンプレート名 = OraDr.GetString(2)
                    If OraDr.IsDBNull(3) = False Then gudt帳票管理情報.str処理関数 = OraDr.GetString(3)
                    If OraDr.IsDBNull(4) = False Then gudt帳票管理情報.intプレビューフラグ = OraDr.GetInt32(4)
                    If OraDr.IsDBNull(5) = False Then gudt帳票管理情報.int出力形式区分 = OraDr.GetInt32(5)
                    If OraDr.IsDBNull(6) = False Then gudt帳票管理情報.strシート名１ = OraDr.GetString(6)
                    If OraDr.IsDBNull(7) = False Then gudt帳票管理情報.int最大明細行数１ = OraDr.GetInt32(7)
                    If OraDr.IsDBNull(8) = False Then gudt帳票管理情報.int明細間隔行数１ = OraDr.GetInt32(8)
                    If OraDr.IsDBNull(9) = False Then gudt帳票管理情報.strシート名２ = OraDr.GetString(9)
                    If OraDr.IsDBNull(10) = False Then gudt帳票管理情報.int最大明細行数２ = OraDr.GetInt32(10)
                    If OraDr.IsDBNull(11) = False Then gudt帳票管理情報.int明細間隔行数２ = OraDr.GetInt32(11)
                    If OraDr.IsDBNull(12) = False Then gudt帳票管理情報.strシート名３ = OraDr.GetString(12)
                    If OraDr.IsDBNull(13) = False Then gudt帳票管理情報.int最大明細行数３ = OraDr.GetInt32(13)
                    If OraDr.IsDBNull(14) = False Then gudt帳票管理情報.int明細間隔行数３ = OraDr.GetInt32(14)
                    If OraDr.IsDBNull(15) = False Then gudt帳票管理情報.strシート名４ = OraDr.GetString(15)
                    If OraDr.IsDBNull(16) = False Then gudt帳票管理情報.int最大明細行数４ = OraDr.GetInt32(16)
                    If OraDr.IsDBNull(17) = False Then gudt帳票管理情報.int明細間隔行数４ = OraDr.GetInt32(17)
                    If OraDr.IsDBNull(18) = False Then gudt帳票管理情報.intバーコード種類 = OraDr.GetInt32(18)
                    If OraDr.IsDBNull(19) = False Then gudt帳票管理情報.intバーコード高さ = OraDr.GetInt32(19)
                    If OraDr.IsDBNull(20) = False Then gudt帳票管理情報.intバーコード幅 = OraDr.GetInt32(20)
                    If OraDr.IsDBNull(21) = False Then gudt帳票管理情報.int表示倍率 = OraDr.GetInt32(21)
                    If OraDr.IsDBNull(22) = False Then gudt帳票管理情報.str概要 = OraDr.GetString(22)
                    If OraDr.IsDBNull(23) = False Then gudt帳票管理情報.str備考 = OraDr.GetString(23)

                    DLTP0996S_PROC0001 = PO_intSQLCODE

                    Exit Function

                Case 1

                    DLTP0996S_PROC0001 = PO_intSQLCODE

                    Exit Function

                Case Else

                    log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                    Exit Function

            End Select

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function

    ''' <summary>
    ''' 有効期限切迫データ出力結果を元にテーブルを更新
    ''' </summary>
    ''' <param name="rdocDgv">明細</param>
    ''' <param name="PO_intSQLCODE">Oracleエラーコード（戻値）</param>
    ''' <param name="PO_strSQLERRM">Oracleエラーメッセージ（戻値）</param>
    ''' <returns>True -- 正常終了 False -- エラー</returns>
    Public Shared Function DLTP0201_PROC0024(ByVal rdocDgv As DataGridView,
                                             ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim i As Integer
        Dim intCnt As Integer
        Dim intPrintChk As Integer
        Dim ID() As Integer = Nothing
        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter

        Try

            DLTP0201_PROC0024 = False

            'ストアドプロシージャ設定
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0201.PROC0024"
            Oracmd.CommandType = CommandType.StoredProcedure

            intCnt = 0

            For i = 0 To rdocDgv.RowCount - 1

                If CInt(rdocDgv.Rows(i).Cells(21).Value) = 9 Then
                    Continue For
                End If

                '印刷チェックボックス
                If IsDBNull(rdocDgv.Rows(i).Cells(24).Value) = True Then
                    intPrintChk = 0
                Else
                    intPrintChk = CInt(rdocDgv.Rows(i).Cells(24).Value)
                End If

                '印刷チェックボックスOFFは対象外
                If intPrintChk = 0 Then
                    Continue For
                End If

                ReDim Preserve ID(intCnt)
                ID(intCnt) = CInt(rdocDgv.Rows(i).Cells(23).Value)
                intCnt += 1
            Next

            If intCnt = 0 Then Return True

            OraTran = Oracomm.BeginTransaction

            'インプットパラメータ設定
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int32, ParameterDirection.Input)

            'Outputパラメータ設定
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.Int32, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            PI_01.CollectionType = OracleCollectionType.PLSQLAssociativeArray
            PI_01.Size = intCnt

            'インプット値設定
            PI_01.Value = ID
            PI_02.Value = intCnt

            'ストアドプロシージャCall
            Oracmd.ExecuteNonQuery()

            PO_intSQLCODE = CType(PO_01.Value.ToString, Integer)
            PO_strSQLERRM = PO_02.Value.ToString

            'リターンコードでの処理振り分け
            If PO_intSQLCODE = 0 Then

                OraTran.Commit()

                DLTP0201_PROC0024 = True

                Exit Function

            End If

            log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

            OraTran.Rollback()

        Catch Oraex As OracleException

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            If Not IsNothing(OraTran) Then
                OraTran.Dispose()
            End If

            Oracmd.Dispose()

        End Try

    End Function

End Class
