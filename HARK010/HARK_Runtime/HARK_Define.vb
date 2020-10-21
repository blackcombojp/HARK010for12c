'/*-----------------------------------------------------------------------------
' * COPYRIGHT(C) ITI CORPORATION 2020
' * ITI CONFIDENTIAL AND PROPRIETARY
' *
' * All rights reserved by ITI Corporation.
' *-----------------------------------------------------------------------------/
Option Compare Binary
Option Explicit On
Option Strict On

Module HARK_Define

    ''*********************************************************************
    ''各メッセージ（アップデート）
    ''*********************************************************************
    Public Const MSG_UPD001 As String = "最新版がリリースされていますので更新します"
    Public Const MSG_UPD002 As String = "更新は異常終了しました"
    Public Const MSG_UPD003 As String = "お使いのバージョンが最新版です"

    ''*********************************************************************
    ''各メッセージ（共通）
    ''*********************************************************************
    Public Const MSG_COM001 As String = "表示をクリアしますか？"
    Public Const MSG_COM002 As String = "対象データはありません"
    Public Const MSG_COM003 As String = "ツールを終了しますか？"
    Public Const MSG_COM004 As String = "サーバとの接続が遮断されています"
    Public Const MSG_COM005 As String = "既に起動しています"
    'Public Const MSG_COM006 As String = "PrintScreenは無効です"
    Public Const MSG_COM007 As String = "事業所を選択してください"
    Public Const MSG_COM012 As String = "プログラムを選択してください"
    Public Const MSG_COM013 As String = "得意先を選択してください"
    Public Const MSG_COM014 As String = "需要先を選択してください"
    Public Const MSG_COM015 As String = "対象日を指定してください"
    Public Const MSG_COM016 As String = "お使いのPC設定（解像度）では使用できません"
    Public Const MSG_COM017 As String = "解像度の設定を変更してください"
    Public Const MSG_COM018 As String = "対象日の指定が不正です"
    Public Const MSG_COM019 As String = "商品コードを指定してください"
    Public Const MSG_COM020 As String = "相手先品番を指定してください"
    Public Const MSG_COM021 As String = "件数："
    Public Const MSG_COM022 As String = "得意先コードを指定してください"
    Public Const MSG_COM023 As String = "需要先コードを指定してください"

    Public Const MSG_COM801 As String = "該当のプログラムは印刷設定がありません"
    Public Const MSG_COM802 As String = "印刷はできません"
    Public Const MSG_COM803 As String = "印刷データ作成処理で異常は発生しました"




    Public Const MSG_COM901 As String = "システム管理者までご連絡ください"
    Public Const MSG_COM902 As String = "予期しないエラーが発生しました"
    Public Const MSG_COM903 As String = "汎用データ検索ツール for 物流管理を再起動してください"

    ''*********************************************************************
    ''各システム規定値
    ''*********************************************************************
    Public Const DUMMY_INTCODE As Integer = 999999999    '検索空白時ダミー定数
    Public Const DUMMY_LNGCODE As Long = 9999999999      '検索空白時ダミー定数
    Public Const DUMMY_STRCODE As String = "999999999"   '検索空白時ダミー定数
    Public Const DUMMY_REGKEY As String = "NULL"         'ダミーレジストリキー
    Public Const DUMMY_FILENAME As String = "DUMMY.txt"  'ダミーファイル名
    Public Const DUMMY_DATESTRING As String = "____/__/__"  'ダミー日付


    ''*********************************************************************
    ''Get_OSVersion関数戻値
    ''*********************************************************************
    Public Const OS_WINDOWS95 As Integer = 0
    Public Const OS_WINDOWS98 As Integer = 1
    Public Const OS_WINDOWSME As Integer = 2
    Public Const OS_WINDOWSNT3 As Integer = 3
    Public Const OS_WINDOWSNT31 As Integer = 4
    Public Const OS_WINDOWSNT35 As Integer = 5
    Public Const OS_WINDOWSNT351 As Integer = 6
    Public Const OS_WINDOWSNT4 As Integer = 7
    Public Const OS_WINDOWS2000 As Integer = 8
    Public Const OS_WINDOWSXP As Integer = 9
    Public Const OS_WINDOWSSERVER2003 As Integer = 10
    Public Const OS_WINDOWSVISTA As Integer = 11
    Public Const OS_WINDOWS7 As Integer = 12
    Public Const OS_WINDOWS32s As Integer = 13
    Public Const OS_WINDOWSCE As Integer = 14
    Public Const OS_UNIX As Integer = 15
    Public Const OS_XBOX As Integer = 16
    Public Const OS_MACINTOSH As Integer = 17
    Public Const OS_UNKNOWN As Integer = 18
    Public Const OS_WINDOWS8 As Integer = 19
    Public Const OS_WINDOWS81 As Integer = 20
    Public Const OS_WINDOWS10 As Integer = 21

    ''*********************************************************************
    ''Entry_Check関数用定数(Check_SIZE)
    ''*********************************************************************
    'Public Const CHECK_SIZE_WIDE As Integer = 1           '全角
    'Public Const CHECK_SIZE_NARROW As Integer = 2         '半角
    'Public Const CHECK_SIZE_BOTH As Integer = 0           '共用

    ''*********************************************************************
    ''Entry_Check関数用定数(Check_STYLE)
    ''*********************************************************************
    'Public Const CHECK_STYLE_NUMBER As Integer = 0        '数字のみ
    'Public Const CHECK_STYLE_ALPH As Integer = 1          '英数字のみ
    'Public Const CHECK_STYLE_ELSE As Integer = 2          'その他

    ''*********************************************************************
    ''Entry_Check関数用定数(Check_LEN)
    ''*********************************************************************
    'Public Const CHECK_LEN_MAKERCODE As Integer = 10            'メーカコード
    'Public Const CHECK_LEN_ITEMMAKERCODE As Integer = 20        'メーカ品番
    'Public Const CHECK_LEN_HPITEMOCDE As Integer = 30           '院内コード

End Module
