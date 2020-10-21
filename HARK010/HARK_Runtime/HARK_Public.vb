'/*-----------------------------------------------------------------------------
' * COPYRIGHT(C) ITI CORPORATION 2020
' * ITI CONFIDENTIAL AND PROPRIETARY
' *
' * All rights reserved by ITI Corporation.
' *-----------------------------------------------------------------------------/
Option Compare Binary
Option Explicit On
Option Strict On

Module HARK_Public

    '*********************************************************************
    '各ファイルパス
    '*********************************************************************
    Public gstrAppFilePath As String            '自exeパス
    Public gstrApplicationDataPath As String    'カレントユーザーApplicationDataパス
    Public gstrlogFilePath As String            '各種ログファイルパス
    Public gstrLogFileName As String            'ログファイルパス
    Public gstrExecuteLogFileName As String     '処理実行ログファイルパス

    '*********************************************************************
    '各戻値
    '*********************************************************************
    Public gintMsg As Integer            'Msgbox戻値
    Public gblRtn As Boolean             'Bool型戻値
    Public gintRtn As Integer            'int型戻値
    Public gstrDate As String            '日付型戻値
    Public gstrRtn As String             '文字列型戻値
    Public gintSQLCODE As Integer        'Oracleエラーコード
    Public gstrSQLERRM As String         'Oracleエラーメッセージ
    Public gstrセッション端末名 As String   'セッション情報(端末)
    Public gintセッションID As Integer      'セッション情報(ID)


    '*********************************************************************
    'ソリューション変数
    '*********************************************************************
    Public gstr部門名 As String                     '部門名


    ''*********************************************************************
    ''レコードカウント変数
    ''*********************************************************************
    Public gintResultCnt As Integer           '検索結果件数
    Public gint得意先Cnt As Integer           '取得得意先数
    Public gint需要先Cnt As Integer           '取得需要先数
    Public gint事業所Cnt As Integer           '取得事業所数
    Public gintサブプログラムCnt As Integer   '取得サブプログラム数

End Module
