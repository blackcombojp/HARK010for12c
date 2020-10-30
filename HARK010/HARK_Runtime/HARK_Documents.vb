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
Imports HARK010.HARK_Common
Imports HARK010.HARK_Sub
Imports HARK010.HARK_DBCommon

Public Class HARK_Documents

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　帳票データ作成【出荷検品未完了一覧】
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　rdocDoc -- 帳票データ
    ' *　戻値　　　　　　：　True -- 正常終了,False -- 異常終了
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTN0001_CreateDocument(ByVal rdocDgv As DataGridView, ByRef rdocReport As CellReport, ByRef rdocDoc As Document) As Boolean

        Dim intCnt As Integer
        Dim intRowCnt As Integer
        Dim intDetaRowCnt As Integer
        Dim blnRecFLG As Boolean = False
        Dim dbRowHeight As Double

        Try

            DTN0001_CreateDocument = False

            rdocReport.FileName = gstrAppFilePath & "\rpt\" & gudt帳票管理情報.strテンプレート名
            rdocReport.Report.Start()
            rdocReport.Report.File()

            intRowCnt = 0
            intCnt = 0
            intDetaRowCnt = 1

            rdocReport.Page.Start(gudt帳票管理情報.strシート名１, "1-99999")

            For i = 0 To rdocDgv.RowCount - 1

                If intDetaRowCnt > gudt帳票管理情報.int最大明細行数１ Then

                    rdocReport.Page.End()
                    rdocReport.Page.Start(gudt帳票管理情報.strシート名１, "1-99999")
                    intRowCnt = 0
                    intDetaRowCnt = 1

                End If

                If blnRecFLG = False Then
                    '最初の１回のみ処理する
                    dbRowHeight = rdocReport.Cell("**PickingNo", 0, intRowCnt).RowHeight
                End If

                '行高さ調整
                rdocReport.Cell("**PickingNo", 0, intRowCnt).RowHeight = dbRowHeight

                If IsNull(rdocDgv.Rows(intCnt).Cells(0).Value.ToString) = False Then rdocReport.Cell("**PickingNo", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(0).Value) & "-" & CStr(rdocDgv.Rows(intCnt).Cells(1).Value) 'ピッキング番号
                'If IsNull(Dgv.Rows(intCnt).Cells(1).Value.ToString) = False Then CellReport.Cell("**GyoNo", 0, intRowCnt).Value = Dgv.Rows(intCnt).Cells(1).Value 'ピッキング行番号
                If IsNull(rdocDgv.Rows(intCnt).Cells(2).Value.ToString) = False Then rdocReport.Cell("**JyuyosakiMei", 0, intRowCnt).Value = rdocDgv.Rows(intCnt).Cells(2).Value '需要先名
                If IsNull(rdocDgv.Rows(intCnt).Cells(3).Value.ToString) = False Then rdocReport.Cell("**JyuyobusyoMei", 0, intRowCnt).Value = rdocDgv.Rows(intCnt).Cells(3).Value '需要先部署名
                If IsNull(rdocDgv.Rows(intCnt).Cells(4).Value.ToString) = False Then rdocReport.Cell("**ItemCode", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(4).Value) '商品コード
                If IsNull(rdocDgv.Rows(intCnt).Cells(5).Value.ToString) = False Then rdocReport.Cell("**MakerName", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(5).Value) 'メーカ名
                If IsNull(rdocDgv.Rows(intCnt).Cells(6).Value.ToString) = False Then rdocReport.Cell("**MakerItemCode", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(6).Value) 'メーカ品番
                If IsNull(rdocDgv.Rows(intCnt).Cells(7).Value.ToString) = False Then rdocReport.Cell("**ItemName", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(7).Value) '商品名
                If IsNull(rdocDgv.Rows(intCnt).Cells(8).Value.ToString) = False Then rdocReport.Cell("**Standard", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(8).Value) '規格
                If IsNull(rdocDgv.Rows(intCnt).Cells(9).Value.ToString) = False Then rdocReport.Cell("**SyukaSuryo", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(9).Value) '出荷数量
                If IsNull(rdocDgv.Rows(intCnt).Cells(10).Value.ToString) = False Then rdocReport.Cell("**Tani", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(10).Value) '単位名
                If IsNull(rdocDgv.Rows(intCnt).Cells(11).Value.ToString) = False Then rdocReport.Cell("**OrderNo", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(11).Value) '顧客注文番号
                If IsNull(rdocDgv.Rows(intCnt).Cells(12).Value.ToString) = False Then rdocReport.Cell("**PickingPrintDate", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(12).Value) 'ピッキングリスト印刷日
                If IsNull(rdocDgv.Rows(intCnt).Cells(13).Value.ToString) = False Then rdocReport.Cell("**Jyutyukeitai", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(13).Value) '受注形態

                intRowCnt += gudt帳票管理情報.int明細間隔行数１

                intCnt += 1
                intDetaRowCnt += 1

                blnRecFLG = True

            Next

            rdocReport.Page.End()
            rdocReport.Report.End()
            rdocDoc = rdocReport.Document

            DTN0001_CreateDocument = True

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　帳票データ作成【PHsmos受注一覧】
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　rdocDoc -- 帳票データ
    ' *　戻値　　　　　　：　True -- 正常終了,False -- 異常終了
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTN0006_CreateDocument(ByVal rdocDgv As DataGridView, ByRef rdocReport As CellReport, ByRef rdocDoc As Document) As Boolean

        Dim intCnt As Integer
        Dim intRowCnt As Integer
        Dim intDetaRowCnt As Integer
        Dim blnRecFLG As Boolean = False
        Dim dbRowHeight As Double

        Try

            DTN0006_CreateDocument = False

            rdocReport.FileName = gstrAppFilePath & "\rpt\" & gudt帳票管理情報.strテンプレート名
            rdocReport.Report.Start()
            rdocReport.Report.File()

            intRowCnt = 0
            intCnt = 0
            intDetaRowCnt = 1

            rdocReport.Page.Start(gudt帳票管理情報.strシート名１, "1-99999")

            For i = 0 To rdocDgv.RowCount - 1

                If intDetaRowCnt > gudt帳票管理情報.int最大明細行数１ Then

                    rdocReport.Page.End()
                    rdocReport.Page.Start(gudt帳票管理情報.strシート名１, "1-99999")
                    intRowCnt = 0
                    intDetaRowCnt = 1

                End If

                If blnRecFLG = False Then
                    '最初の１回のみ処理する
                    dbRowHeight = rdocReport.Cell("**HatyuNo", 0, intRowCnt).RowHeight
                End If

                '行高さ調整
                rdocReport.Cell("**HatyuNo", 0, intRowCnt).RowHeight = dbRowHeight

                If IsNull(rdocDgv.Rows(intCnt).Cells(0).Value.ToString) = False Then rdocReport.Cell("**HatyuNo", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(0).Value) & "-" & CStr(rdocDgv.Rows(intCnt).Cells(1).Value)
                'If IsNull(Dgv.Rows(intCnt).Cells(1).Value.ToString) = False Then CellReport.Cell("**GyoNo", 0, intRowCnt).Value = Dgv.Rows(intCnt).Cells(1).Value 
                If IsNull(rdocDgv.Rows(intCnt).Cells(2).Value.ToString) = False Then rdocReport.Cell("**HatyuDate", 0, intRowCnt).Value = rdocDgv.Rows(intCnt).Cells(2).Value
                If IsNull(rdocDgv.Rows(intCnt).Cells(4).Value.ToString) = False Then rdocReport.Cell("**Tokuisakimei", 0, intRowCnt).Value = rdocDgv.Rows(intCnt).Cells(4).Value
                If IsNull(rdocDgv.Rows(intCnt).Cells(5).Value.ToString) = False Then rdocReport.Cell("**JyuyosakiMei", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(5).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(6).Value.ToString) = False Then rdocReport.Cell("**JyuyobusyoMei", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(6).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(7).Value.ToString) = False Then rdocReport.Cell("**ItemCode", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(7).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(9).Value.ToString) = False Then rdocReport.Cell("**MakerName", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(9).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(10).Value.ToString) = False Then rdocReport.Cell("**MakerItemCode", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(10).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(11).Value.ToString) = False Then rdocReport.Cell("**ItemName", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(11).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(12).Value.ToString) = False Then rdocReport.Cell("**Standard", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(12).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(13).Value.ToString) = False Then rdocReport.Cell("**HatyuSuryo", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(13).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(14).Value.ToString) = False Then rdocReport.Cell("**Tani", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(14).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(15).Value.ToString) = False Then rdocReport.Cell("**Jyutyukeitai", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(15).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(16).Value.ToString) = False Then rdocReport.Cell("**Error", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(16).Value)

                intRowCnt += gudt帳票管理情報.int明細間隔行数１

                intCnt += 1
                intDetaRowCnt += 1

                blnRecFLG = True

            Next

            rdocReport.Page.End()
            rdocReport.Report.End()
            rdocDoc = rdocReport.Document

            DTN0006_CreateDocument = True

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try


    End Function
    '/*-----------------------------------------------------------------------------
    ' *　モジュール機能　：　帳票データ作成【有効期限切迫情報】
    ' *
    ' *　注意、制限事項　：　なし
    ' *　引数１　　　　　：　rdocDoc -- 帳票データ
    ' *　戻値　　　　　　：　True -- 正常終了,False -- 異常終了
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTN0009_CreateDocument(ByVal rdocDgv As DataGridView, ByRef rdocReport As CellReport, ByRef rdocDoc As Document) As Boolean

        Dim intCnt As Integer
        Dim intRowCnt As Integer
        Dim intDetaRowCnt As Integer
        Dim blnRecFLG As Boolean = False
        Dim dbRowHeight As Double

        Try

            DTN0009_CreateDocument = False

            rdocReport.FileName = gstrAppFilePath & "\rpt\" & gudt帳票管理情報.strテンプレート名
            rdocReport.Report.Start()
            rdocReport.Report.File()

            intRowCnt = 0
            intCnt = 0
            intDetaRowCnt = 1

            rdocReport.Page.Start(gudt帳票管理情報.strシート名１, "1-99999")

            For i = 0 To rdocDgv.RowCount - 1

                '自社貸出は対象外
                If CInt(rdocDgv.Rows(i).Cells(21).Value) = 9 Or CInt(rdocDgv.Rows(i).Cells(21).Value) = 1 Then
                    intCnt += 1
                    Continue For
                End If

                If intDetaRowCnt > gudt帳票管理情報.int最大明細行数１ Then

                    rdocReport.Page.End()
                    rdocReport.Page.Start(gudt帳票管理情報.strシート名１, "1-99999")
                    intRowCnt = 0
                    intDetaRowCnt = 1

                End If

                If blnRecFLG = False Then
                    '最初の１回のみ処理する
                    dbRowHeight = rdocReport.Cell("**JyuyosakiName", 0, intRowCnt).RowHeight
                End If

                '行高さ調整
                rdocReport.Cell("**JyuyosakiName", 0, intRowCnt).RowHeight = dbRowHeight

                If IsNull(rdocDgv.Rows(intCnt).Cells(1).Value.ToString) = False Then rdocReport.Cell("**NohinNo", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(1).Value) & "-" & CStr(rdocDgv.Rows(intCnt).Cells(2).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(22).Value.ToString) = False Then rdocReport.Cell("**SeikyuDate", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(22).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(4).Value.ToString) = False Then rdocReport.Cell("**JyuyosakiName", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(4).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(5).Value.ToString) = False Then rdocReport.Cell("**JyuyoBusyoName", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(5).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(7).Value.ToString) = False Then rdocReport.Cell("**MakerName", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(7).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(9).Value.ToString) = False Then rdocReport.Cell("**ShohinName", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(9).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(8).Value.ToString) = False Then rdocReport.Cell("**MakerHinban", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(8).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(12).Value.ToString) = False Then rdocReport.Cell("**Suryo", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(12).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(13).Value.ToString) = False Then rdocReport.Cell("**Tani", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(13).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(18).Value.ToString) = False Then rdocReport.Cell("**YukokigenDate", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(18).Value)
                If IsNull(rdocDgv.Rows(intCnt).Cells(14).Value.ToString) = False Then rdocReport.Cell("**KenpinDate", 0, intRowCnt).Value = CStr(rdocDgv.Rows(intCnt).Cells(14).Value)

                intRowCnt += gudt帳票管理情報.int明細間隔行数１

                intCnt += 1
                intDetaRowCnt += 1

                blnRecFLG = True

            Next

            rdocReport.Page.End()
            rdocReport.Report.End()
            rdocDoc = rdocReport.Document

            '印刷済みフラグON
            gblRtn = DLTP0201_PROC0024(rdocDgv, gintSQLCODE, gstrSQLERRM)

            DTN0009_CreateDocument = True

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try


    End Function

End Class
