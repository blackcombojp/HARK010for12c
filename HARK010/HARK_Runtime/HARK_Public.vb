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
    '�e�t�@�C���p�X
    '*********************************************************************
    Public gstrAppFilePath As String            '��exe�p�X
    Public gstrApplicationDataPath As String    '�J�����g���[�U�[ApplicationData�p�X
    Public gstrlogFilePath As String            '�e�탍�O�t�@�C���p�X
    Public gstrLogFileName As String            '���O�t�@�C���p�X
    Public gstrExecuteLogFileName As String     '�������s���O�t�@�C���p�X

    '*********************************************************************
    '�e�ߒl
    '*********************************************************************
    Public gintMsg As Integer            'Msgbox�ߒl
    Public gblRtn As Boolean             'Bool�^�ߒl
    Public gintRtn As Integer            'int�^�ߒl
    Public gstrDate As String            '���t�^�ߒl
    Public gstrRtn As String             '������^�ߒl
    Public gintSQLCODE As Integer        'Oracle�G���[�R�[�h
    Public gstrSQLERRM As String         'Oracle�G���[���b�Z�[�W
    Public gstr�Z�b�V�����[���� As String   '�Z�b�V�������(�[��)
    Public gint�Z�b�V����ID As Integer      '�Z�b�V�������(ID)


    '*********************************************************************
    '�\�����[�V�����ϐ�
    '*********************************************************************
    Public gstr���喼 As String                     '���喼


    ''*********************************************************************
    ''���R�[�h�J�E���g�ϐ�
    ''*********************************************************************
    Public gintResultCnt As Integer           '�������ʌ���
    Public gint���Ӑ�Cnt As Integer           '�擾���Ӑ搔
    Public gint���v��Cnt As Integer           '�擾���v�搔
    Public gint���Ə�Cnt As Integer           '�擾���Ə���
    Public gint�T�u�v���O����Cnt As Integer   '�擾�T�u�v���O������

End Module
