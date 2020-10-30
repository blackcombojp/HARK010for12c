'/*-----------------------------------------------------------------------------
' * COPYRIGHT(C) ITI CORPORATION 2020
' * ITI CONFIDENTIAL AND PROPRIETARY
' *
' * All rights reserved by ITI Corporation.
' *-----------------------------------------------------------------------------/
Option Compare Binary
Option Explicit On
Option Strict On


Imports HARK010.HARK_Common
Imports HARK010.HARK_DBCommon
Public Class HARK_Sub
    Private Shared _mutex As Threading.Mutex
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    '/*-----------------------------------------------------------------------------
    ' *  モジュール機能： メイン画面起動前処理
    ' *
    ' *  注意、制限事項  ：なし
    ' *  引数　　　　　　：なし
    ' *　戻値　　　　　　：なし
    ' *-----------------------------------------------------------------------------/
    Public Shared Sub Main()

        'TLS1.2のみ許可(Webアクセス)
        Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12

        Dim hasHandle As Boolean = False

        Try

            log4net.NDC.Push(My.Application.Info.Version.ToString)

            _mutex = New Threading.Mutex(False, My.Application.Info.ProductName)

            '自exeパス取得
            gstrAppFilePath = Get_AppPath()

            'カレントユーザーApplicationDataパス取得
            gstrApplicationDataPath = Set_FilePath(Get_ApplicationPath(), "HARK")

            'ログファイルパス取得
            gstrlogFilePath = Set_FilePath(gstrApplicationDataPath, "log")

            '各ファイルパス設定
            gstrLogFileName = Set_FilePath(gstrlogFilePath, "HARK010Err.Log")
            gstrExecuteLogFileName = Set_FilePath(gstrlogFilePath, "HARK010Execute.Log")

            Try
                'ミューテックスの所有権を要求する
                hasHandle = _mutex.WaitOne(0, False)

            Catch ex As Threading.AbandonedMutexException
                hasHandle = True
            End Try

            If hasHandle = False Then
                log.Error(Set_ErrMSG(0, MSG_COM005))
                MsgBox(MSG_COM005, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, My.Application.Info.Title)
                Return
            End If

            'Oracle接続
            If OraConnect() = False Then
                log.Error(Set_ErrMSG(0, MSG_COM004))
                MsgBox(MSG_COM004, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, My.Application.Info.Title)
                Return
            End If

            '部門名取得
            If My.Settings.事業所コード <> 0 Then
                If DLTP0900_PROC0002("Sub_Main", gintSQLCODE, gstrSQLERRM) = False Then
                    MsgBox(gintSQLCODE & "-" & gstrSQLERRM & vbCr & MSG_COM902 & vbCr & MSG_COM901, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                    log.Error(Set_ErrMSG(gintSQLCODE, gstrSQLERRM))
                    Application.Exit()
                    Return
                End If
            End If

            '事業所一覧取得
            If DLTP0901_PROC0001(gintSQLCODE, gstrSQLERRM) = False Then
                MsgBox(gintSQLCODE & "-" & gstrSQLERRM & vbCr & MSG_COM902 & vbCr & MSG_COM901, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                log.Error(Set_ErrMSG(gintSQLCODE, gstrSQLERRM))
                Application.Exit()
                Return
            End If

            'アプリ起動
            Application.Run(New HARK001)

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

            Application.Exit()

        Finally

            If hasHandle Then
                _mutex.ReleaseMutex()
            End If

            _mutex.Close()

        End Try

    End Sub
    '/*-----------------------------------------------------------------------------
    ' *  モジュール機能： 指定フォームIDのキャプションを設定
    ' *
    ' *  注意、制限事項  ：なし
    ' *  引数　　　　　　：FormID ・・各フォーム名（Define定義）
    ' *  　　　　　　　　：SenderName ・・各フォームオブジェクト名
    ' *　戻値　　　　　　：関数・・フォームキャプション
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Set_FormTitle(ByVal ProgramID As String,
                                         ByVal FormID As String) As String

        Try
            'パスにファイル名を追加
            Set_FormTitle = ProgramID & " " & FormID & " 【" & My.Application.Info.CompanyName & "】"

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
#Region "エラーメッセージ整形 Set_ErrMSG"
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　エラーメッセージ整形
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　strErrCode    -- エラーコード
    ' *　引数２　　　　　：　strErrMessage -- エラーメッセージ
    ' *　戻値　　　　　　：　整形後エラーメッセージ
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Set_ErrMSG(ByVal strErrCode As Integer, ByVal strErrMessage As String) As String

        Dim strBuff As String

        strBuff = CType(strErrCode, String) & " " & strErrMessage

        Set_ErrMSG = strBuff

    End Function
#End Region

End Class
