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
    ''�e���b�Z�[�W�i�A�b�v�f�[�g�j
    ''*********************************************************************
    Public Const MSG_UPD001 As String = "�ŐV�ł������[�X����Ă��܂��̂ōX�V���܂�"
    Public Const MSG_UPD002 As String = "�X�V�ُ͈�I�����܂���"
    Public Const MSG_UPD003 As String = "���g���̃o�[�W�������ŐV�łł�"

    ''*********************************************************************
    ''�e���b�Z�[�W�i���ʁj
    ''*********************************************************************
    Public Const MSG_COM001 As String = "�\�����N���A���܂����H"
    Public Const MSG_COM002 As String = "�Ώۃf�[�^�͂���܂���"
    Public Const MSG_COM003 As String = "�c�[�����I�����܂����H"
    Public Const MSG_COM004 As String = "�T�[�o�Ƃ̐ڑ����Ւf����Ă��܂�"
    Public Const MSG_COM005 As String = "���ɋN�����Ă��܂�"
    'Public Const MSG_COM006 As String = "PrintScreen�͖����ł�"
    Public Const MSG_COM007 As String = "���Ə���I�����Ă�������"
    Public Const MSG_COM012 As String = "�v���O������I�����Ă�������"
    Public Const MSG_COM013 As String = "���Ӑ��I�����Ă�������"
    Public Const MSG_COM014 As String = "���v���I�����Ă�������"
    Public Const MSG_COM015 As String = "�Ώۓ����w�肵�Ă�������"
    Public Const MSG_COM016 As String = "���g����PC�ݒ�i�𑜓x�j�ł͎g�p�ł��܂���"
    Public Const MSG_COM017 As String = "�𑜓x�̐ݒ��ύX���Ă�������"
    Public Const MSG_COM018 As String = "�Ώۓ��̎w�肪�s���ł�"
    Public Const MSG_COM019 As String = "���i�R�[�h���w�肵�Ă�������"
    Public Const MSG_COM020 As String = "�����i�Ԃ��w�肵�Ă�������"
    Public Const MSG_COM021 As String = "�����F"
    Public Const MSG_COM022 As String = "���Ӑ�R�[�h���w�肵�Ă�������"
    Public Const MSG_COM023 As String = "���v��R�[�h���w�肵�Ă�������"

    Public Const MSG_COM801 As String = "�Y���̃v���O�����͈���ݒ肪����܂���"
    Public Const MSG_COM802 As String = "����͂ł��܂���"
    Public Const MSG_COM803 As String = "����f�[�^�쐬�����ňُ�͔������܂���"




    Public Const MSG_COM901 As String = "�V�X�e���Ǘ��҂܂ł��A����������"
    Public Const MSG_COM902 As String = "�\�����Ȃ��G���[���������܂���"
    Public Const MSG_COM903 As String = "�ėp�f�[�^�����c�[�� for �����Ǘ����ċN�����Ă�������"

    ''*********************************************************************
    ''�e�V�X�e���K��l
    ''*********************************************************************
    Public Const DUMMY_INTCODE As Integer = 999999999    '�����󔒎��_�~�[�萔
    Public Const DUMMY_LNGCODE As Long = 9999999999      '�����󔒎��_�~�[�萔
    Public Const DUMMY_STRCODE As String = "999999999"   '�����󔒎��_�~�[�萔
    Public Const DUMMY_REGKEY As String = "NULL"         '�_�~�[���W�X�g���L�[
    Public Const DUMMY_FILENAME As String = "DUMMY.txt"  '�_�~�[�t�@�C����
    Public Const DUMMY_DATESTRING As String = "____/__/__"  '�_�~�[���t


    ''*********************************************************************
    ''Get_OSVersion�֐��ߒl
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
    ''Entry_Check�֐��p�萔(Check_SIZE)
    ''*********************************************************************
    'Public Const CHECK_SIZE_WIDE As Integer = 1           '�S�p
    'Public Const CHECK_SIZE_NARROW As Integer = 2         '���p
    'Public Const CHECK_SIZE_BOTH As Integer = 0           '���p

    ''*********************************************************************
    ''Entry_Check�֐��p�萔(Check_STYLE)
    ''*********************************************************************
    'Public Const CHECK_STYLE_NUMBER As Integer = 0        '�����̂�
    'Public Const CHECK_STYLE_ALPH As Integer = 1          '�p�����̂�
    'Public Const CHECK_STYLE_ELSE As Integer = 2          '���̑�

    ''*********************************************************************
    ''Entry_Check�֐��p�萔(Check_LEN)
    ''*********************************************************************
    'Public Const CHECK_LEN_MAKERCODE As Integer = 10            '���[�J�R�[�h
    'Public Const CHECK_LEN_ITEMMAKERCODE As Integer = 20        '���[�J�i��
    'Public Const CHECK_LEN_HPITEMOCDE As Integer = 30           '�@���R�[�h

End Module
