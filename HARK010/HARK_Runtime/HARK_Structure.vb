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


    'TuvOê
    Public Structure TuvOê

        Public strTuvOR[h As String    'TuvOR[h
        Public strTuvO¼ As String        'TuvO¼

        Public Overrides Function ToString() As String
            Return strTuvO¼
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As String)
            strTuvO¼ = Name
            strTuvOR[h = CD
        End Sub

    End Structure


    '¾Óæê
    Public Structure ¾Óæê

        Public lng¾ÓæR[h As Long      '¾ÓæR[h
        Public str¾Óæ¼ As String        '¾Óæ¼

        Public Overrides Function ToString() As String
            Return str¾Óæ¼
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As Long)
            str¾Óæ¼ = Name
            lng¾ÓæR[h = CD
        End Sub

    End Structure


    'ùvæê
    Public Structure ùvæê

        Public lngùvæR[h As Long      'ùvæR[h
        Public strùvæ¼ As String        'ùvæ¼

        Public Overrides Function ToString() As String
            Return strùvæ¼
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As Long)
            strùvæ¼ = Name
            lngùvæR[h = CD
        End Sub

    End Structure


    'Æê
    Public Structure Æê

        Public intÆR[h As Integer   'ÆR[h
        Public strÆ¼ As String        'Æ¼

        Public Overrides Function ToString() As String
            Return strÆ¼
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As Integer)
            strÆ¼ = Name
            intÆR[h = CD
        End Sub

    End Structure

    'vO}X^îñ
    Public Structure Struc_vO}X^

        Dim strÖ As String
        Dim stroÍwb_ As String
        Dim stroÍæØ¶ As String
        Dim intõðP As Integer
        Dim intõðQ As Integer
        Dim intõðR As Integer
        Dim intõðS As Integer
        Dim intõðT As Integer
        Dim intõðU As Integer
        Dim intõðV As Integer
        Dim intõðW As Integer
        Dim intõðX As Integer
        Dim intõðPO As Integer
        Dim strõðPqg As String
        Dim strõðQqg As String
        Dim strõðRqg As String
        Dim strõðSqg As String
        Dim strõðTqg As String
        Dim strõðUqg As String
        Dim strõðVqg As String
        Dim strõðWqg As String
        Dim strõðXqg As String
        Dim strõðPOqg As String
        Dim intTuvO_ID As Integer
        Public Sub IsClear()

            strÖ = vbNullString
            stroÍwb_ = vbNullString
            stroÍæØ¶ = vbNullString
            intõðP = 0
            intõðQ = 0
            intõðR = 0
            intõðS = 0
            intõðT = 0
            intõðU = 0
            intõðV = 0
            intõðW = 0
            intõðX = 0
            intõðPO = 0
            strõðPqg = vbNullString
            strõðQqg = vbNullString
            strõðRqg = vbNullString
            strõðSqg = vbNullString
            strõðTqg = vbNullString
            strõðUqg = vbNullString
            strõðVqg = vbNullString
            strõðWqg = vbNullString
            strõðXqg = vbNullString
            strõðPOqg = vbNullString
            intTuvO_ID = 0

        End Sub
    End Structure

    ' [Çîñ
    Public Structure Struc_ [Çîñ

        Dim lng [ÇÔ As Long
        Dim str [¼ As String
        Dim strev[g¼ As String
        Dim strÖ As String
        Dim intvr[tO As Integer
        Dim intoÍ`®æª As Integer
        Dim strV[g¼P As String
        Dim intÅå¾×sP As Integer
        Dim int¾×ÔusP As Integer
        Dim strV[g¼Q As String
        Dim intÅå¾×sQ As Integer
        Dim int¾×ÔusQ As Integer
        Dim strV[g¼R As String
        Dim intÅå¾×sR As Integer
        Dim int¾×ÔusR As Integer
        Dim strV[g¼S As String
        Dim intÅå¾×sS As Integer
        Dim int¾×ÔusS As Integer
        Dim into[R[híÞ As Integer
        Dim into[R[h³ As Integer
        Dim into[R[h As Integer
        Dim int\¦{¦ As Integer
        Dim strTv As String
        Dim strõl As String
        Public Sub IsClear()

            lng [ÇÔ = 0
            str [¼ = vbNullString
            strev[g¼ = vbNullString
            strÖ = vbNullString
            intvr[tO = 0
            intoÍ`®æª = 0
            strV[g¼P = vbNullString
            intÅå¾×sP = 0
            int¾×ÔusP = 0
            strV[g¼Q = vbNullString
            intÅå¾×sQ = 0
            int¾×ÔusQ = 0
            strV[g¼R = vbNullString
            intÅå¾×sR = 0
            int¾×ÔusR = 0
            strV[g¼S = vbNullString
            intÅå¾×sS = 0
            int¾×ÔusS = 0
            into[R[híÞ = 0
            into[R[h³ = 0
            into[R[h = 0
            int\¦{¦ = 0
            strTv = vbNullString
            strõl = vbNullString

        End Sub
    End Structure

    'õÊê
    Public Structure Result
        Public strBuff() As String
    End Structure

    Public TuvOArray() As TuvOê
    Public ¾ÓæArray() As ¾Óæê
    Public ùvæArray() As ùvæê
    Public ÆArray() As Æê
    Public Results() As Result

    Public gudtvO}X^ As Struc_vO}X^
    Public gudt [Çîñ As Struc_ [Çîñ

End Module
