'/*-----------------------------------------------------------------------------
' * COPYRIGHT(C) ITI CORPORATION 2020
' * ITI CONFIDENTIAL AND PROPRIETARY
' *
' * All rights reserved by ITI Corporation.
' *-----------------------------------------------------------------------------/
Option Compare Binary
Option Explicit On
Option Strict On

Module HARK_Structure


    'サブプログラム一覧
    Public Structure サブプログラム一覧

        Public strサブプログラムコード As String    'サブプログラムコード
        Public strサブプログラム名 As String        'サブプログラム名

        Public Overrides Function ToString() As String
            Return strサブプログラム名
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As String)
            strサブプログラム名 = Name
            strサブプログラムコード = CD
        End Sub

    End Structure


    '得意先一覧
    Public Structure 得意先一覧

        Public lng得意先コード As Long      '得意先コード
        Public str得意先名 As String        '得意先名

        Public Overrides Function ToString() As String
            Return str得意先名
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As Long)
            str得意先名 = Name
            lng得意先コード = CD
        End Sub

    End Structure


    '需要先一覧
    Public Structure 需要先一覧

        Public lng需要先コード As Long      '需要先コード
        Public str需要先名 As String        '需要先名

        Public Overrides Function ToString() As String
            Return str需要先名
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As Long)
            str需要先名 = Name
            lng需要先コード = CD
        End Sub

    End Structure


    '事業所一覧
    Public Structure 事業所一覧

        Public int事業所コード As Integer   '事業所コード
        Public str事業所名 As String        '事業所名

        Public Overrides Function ToString() As String
            Return str事業所名
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As Integer)
            str事業所名 = Name
            int事業所コード = CD
        End Sub

    End Structure

    'プログラムマスタ情報
    Public Structure Struc_プログラムマスタ

        Dim str処理関数 As String
        Dim str出力ヘッダ As String
        Dim str出力区切文字 As String
        Dim int検索条件１ As Integer
        Dim int検索条件２ As Integer
        Dim int検索条件３ As Integer
        Dim int検索条件４ As Integer
        Dim int検索条件５ As Integer
        Dim int検索条件６ As Integer
        Dim int検索条件７ As Integer
        Dim int検索条件８ As Integer
        Dim int検索条件９ As Integer
        Dim int検索条件１０ As Integer
        Dim str検索条件１ヒント As String
        Dim str検索条件２ヒント As String
        Dim str検索条件３ヒント As String
        Dim str検索条件４ヒント As String
        Dim str検索条件５ヒント As String
        Dim str検索条件６ヒント As String
        Dim str検索条件７ヒント As String
        Dim str検索条件８ヒント As String
        Dim str検索条件９ヒント As String
        Dim str検索条件１０ヒント As String
        Dim intサブプログラム_ID As Integer
        Public Sub IsClear()

            str処理関数 = vbNullString
            str出力ヘッダ = vbNullString
            str出力区切文字 = vbNullString
            int検索条件１ = 0
            int検索条件２ = 0
            int検索条件３ = 0
            int検索条件４ = 0
            int検索条件５ = 0
            int検索条件６ = 0
            int検索条件７ = 0
            int検索条件８ = 0
            int検索条件９ = 0
            int検索条件１０ = 0
            str検索条件１ヒント = vbNullString
            str検索条件２ヒント = vbNullString
            str検索条件３ヒント = vbNullString
            str検索条件４ヒント = vbNullString
            str検索条件５ヒント = vbNullString
            str検索条件６ヒント = vbNullString
            str検索条件７ヒント = vbNullString
            str検索条件８ヒント = vbNullString
            str検索条件９ヒント = vbNullString
            str検索条件１０ヒント = vbNullString
            intサブプログラム_ID = 0

        End Sub
    End Structure

    '帳票管理情報
    Public Structure Struc_帳票管理情報

        Dim lng帳票管理番号 As Long
        Dim str帳票名 As String
        Dim strテンプレート名 As String
        Dim str処理関数 As String
        Dim intプレビューフラグ As Integer
        Dim int出力形式区分 As Integer
        Dim strシート名１ As String
        Dim int最大明細行数１ As Integer
        Dim int明細間隔行数１ As Integer
        Dim strシート名２ As String
        Dim int最大明細行数２ As Integer
        Dim int明細間隔行数２ As Integer
        Dim strシート名３ As String
        Dim int最大明細行数３ As Integer
        Dim int明細間隔行数３ As Integer
        Dim strシート名４ As String
        Dim int最大明細行数４ As Integer
        Dim int明細間隔行数４ As Integer
        Dim intバーコード種類 As Integer
        Dim intバーコード高さ As Integer
        Dim intバーコード幅 As Integer
        Dim int表示倍率 As Integer
        Dim str概要 As String
        Dim str備考 As String
        Public Sub IsClear()

            lng帳票管理番号 = 0
            str帳票名 = vbNullString
            strテンプレート名 = vbNullString
            str処理関数 = vbNullString
            intプレビューフラグ = 0
            int出力形式区分 = 0
            strシート名１ = vbNullString
            int最大明細行数１ = 0
            int明細間隔行数１ = 0
            strシート名２ = vbNullString
            int最大明細行数２ = 0
            int明細間隔行数２ = 0
            strシート名３ = vbNullString
            int最大明細行数３ = 0
            int明細間隔行数３ = 0
            strシート名４ = vbNullString
            int最大明細行数４ = 0
            int明細間隔行数４ = 0
            intバーコード種類 = 0
            intバーコード高さ = 0
            intバーコード幅 = 0
            int表示倍率 = 0
            str概要 = vbNullString
            str備考 = vbNullString

        End Sub
    End Structure

    '検索結果一覧
    Public Structure Result
        Public strBuff() As String
    End Structure

    Public サブプログラムArray() As サブプログラム一覧
    Public 得意先Array() As 得意先一覧
    Public 需要先Array() As 需要先一覧
    Public 事業所Array() As 事業所一覧
    Public Results() As Result

    Public gudtプログラムマスタ As Struc_プログラムマスタ
    Public gudt帳票管理情報 As Struc_帳票管理情報

End Module
