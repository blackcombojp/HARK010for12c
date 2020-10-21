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


    '�T�u�v���O�����ꗗ
    Public Structure �T�u�v���O�����ꗗ

        Public str�T�u�v���O�����R�[�h As String    '�T�u�v���O�����R�[�h
        Public str�T�u�v���O������ As String        '�T�u�v���O������

        Public Overrides Function ToString() As String
            Return str�T�u�v���O������
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As String)
            str�T�u�v���O������ = Name
            str�T�u�v���O�����R�[�h = CD
        End Sub

    End Structure


    '���Ӑ�ꗗ
    Public Structure ���Ӑ�ꗗ

        Public lng���Ӑ�R�[�h As Long      '���Ӑ�R�[�h
        Public str���Ӑ於 As String        '���Ӑ於

        Public Overrides Function ToString() As String
            Return str���Ӑ於
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As Long)
            str���Ӑ於 = Name
            lng���Ӑ�R�[�h = CD
        End Sub

    End Structure


    '���v��ꗗ
    Public Structure ���v��ꗗ

        Public lng���v��R�[�h As Long      '���v��R�[�h
        Public str���v�於 As String        '���v�於

        Public Overrides Function ToString() As String
            Return str���v�於
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As Long)
            str���v�於 = Name
            lng���v��R�[�h = CD
        End Sub

    End Structure


    '���Ə��ꗗ
    Public Structure ���Ə��ꗗ

        Public int���Ə��R�[�h As Integer   '���Ə��R�[�h
        Public str���Ə��� As String        '���Ə���

        Public Overrides Function ToString() As String
            Return str���Ə���
        End Function

        Public Sub New(ByVal Name As String, ByVal CD As Integer)
            str���Ə��� = Name
            int���Ə��R�[�h = CD
        End Sub

    End Structure

    '�v���O�����}�X�^���
    Public Structure Struc_�v���O�����}�X�^

        Dim str�����֐� As String
        Dim str�o�̓w�b�_ As String
        Dim str�o�͋�ؕ��� As String
        Dim int���������P As Integer
        Dim int���������Q As Integer
        Dim int���������R As Integer
        Dim int���������S As Integer
        Dim int���������T As Integer
        Dim int���������U As Integer
        Dim int���������V As Integer
        Dim int���������W As Integer
        Dim int���������X As Integer
        Dim int���������P�O As Integer
        Dim str���������P�q���g As String
        Dim str���������Q�q���g As String
        Dim str���������R�q���g As String
        Dim str���������S�q���g As String
        Dim str���������T�q���g As String
        Dim str���������U�q���g As String
        Dim str���������V�q���g As String
        Dim str���������W�q���g As String
        Dim str���������X�q���g As String
        Dim str���������P�O�q���g As String
        Dim int�T�u�v���O����_ID As Integer
        Public Sub IsClear()

            str�����֐� = vbNullString
            str�o�̓w�b�_ = vbNullString
            str�o�͋�ؕ��� = vbNullString
            int���������P = 0
            int���������Q = 0
            int���������R = 0
            int���������S = 0
            int���������T = 0
            int���������U = 0
            int���������V = 0
            int���������W = 0
            int���������X = 0
            int���������P�O = 0
            str���������P�q���g = vbNullString
            str���������Q�q���g = vbNullString
            str���������R�q���g = vbNullString
            str���������S�q���g = vbNullString
            str���������T�q���g = vbNullString
            str���������U�q���g = vbNullString
            str���������V�q���g = vbNullString
            str���������W�q���g = vbNullString
            str���������X�q���g = vbNullString
            str���������P�O�q���g = vbNullString
            int�T�u�v���O����_ID = 0

        End Sub
    End Structure

    '���[�Ǘ����
    Public Structure Struc_���[�Ǘ����

        Dim lng���[�Ǘ��ԍ� As Long
        Dim str���[�� As String
        Dim str�e���v���[�g�� As String
        Dim str�����֐� As String
        Dim int�v���r���[�t���O As Integer
        Dim int�o�͌`���敪 As Integer
        Dim str�V�[�g���P As String
        Dim int�ő喾�׍s���P As Integer
        Dim int���׊Ԋu�s���P As Integer
        Dim str�V�[�g���Q As String
        Dim int�ő喾�׍s���Q As Integer
        Dim int���׊Ԋu�s���Q As Integer
        Dim str�V�[�g���R As String
        Dim int�ő喾�׍s���R As Integer
        Dim int���׊Ԋu�s���R As Integer
        Dim str�V�[�g���S As String
        Dim int�ő喾�׍s���S As Integer
        Dim int���׊Ԋu�s���S As Integer
        Dim int�o�[�R�[�h��� As Integer
        Dim int�o�[�R�[�h���� As Integer
        Dim int�o�[�R�[�h�� As Integer
        Dim int�\���{�� As Integer
        Dim str�T�v As String
        Dim str���l As String
        Public Sub IsClear()

            lng���[�Ǘ��ԍ� = 0
            str���[�� = vbNullString
            str�e���v���[�g�� = vbNullString
            str�����֐� = vbNullString
            int�v���r���[�t���O = 0
            int�o�͌`���敪 = 0
            str�V�[�g���P = vbNullString
            int�ő喾�׍s���P = 0
            int���׊Ԋu�s���P = 0
            str�V�[�g���Q = vbNullString
            int�ő喾�׍s���Q = 0
            int���׊Ԋu�s���Q = 0
            str�V�[�g���R = vbNullString
            int�ő喾�׍s���R = 0
            int���׊Ԋu�s���R = 0
            str�V�[�g���S = vbNullString
            int�ő喾�׍s���S = 0
            int���׊Ԋu�s���S = 0
            int�o�[�R�[�h��� = 0
            int�o�[�R�[�h���� = 0
            int�o�[�R�[�h�� = 0
            int�\���{�� = 0
            str�T�v = vbNullString
            str���l = vbNullString

        End Sub
    End Structure

    '�������ʈꗗ
    Public Structure Result
        Public strBuff() As String
    End Structure

    Public �T�u�v���O����Array() As �T�u�v���O�����ꗗ
    Public ���Ӑ�Array() As ���Ӑ�ꗗ
    Public ���v��Array() As ���v��ꗗ
    Public ���Ə�Array() As ���Ə��ꗗ
    Public Results() As Result

    Public gudt�v���O�����}�X�^ As Struc_�v���O�����}�X�^
    Public gudt���[�Ǘ���� As Struc_���[�Ǘ����

End Module
