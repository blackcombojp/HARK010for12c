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
    ' *  ���W���[���@�\�F ���C����ʋN���O����
    ' *
    ' *  ���ӁA��������  �F�Ȃ�
    ' *  �����@�@�@�@�@�@�F�Ȃ�
    ' *�@�ߒl�@�@�@�@�@�@�F�Ȃ�
    ' *-----------------------------------------------------------------------------/
    Public Shared Sub Main()

        'TLS1.2�̂݋���(Web�A�N�Z�X)
        Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12

        Dim hasHandle As Boolean = False

        Try

            log4net.NDC.Push(My.Application.Info.Version.ToString)

            _mutex = New Threading.Mutex(False, My.Application.Info.ProductName)

            '��exe�p�X�擾
            gstrAppFilePath = Get_AppPath()

            '�J�����g���[�U�[ApplicationData�p�X�擾
            gstrApplicationDataPath = Set_FilePath(Get_ApplicationPath(), "HARK")

            '���O�t�@�C���p�X�擾
            gstrlogFilePath = Set_FilePath(gstrApplicationDataPath, "log")

            '�e�t�@�C���p�X�ݒ�
            gstrLogFileName = Set_FilePath(gstrlogFilePath, "HARK010Err.Log")
            gstrExecuteLogFileName = Set_FilePath(gstrlogFilePath, "HARK010Execute.Log")

            Try
                '�~���[�e�b�N�X�̏��L����v������
                hasHandle = _mutex.WaitOne(0, False)

            Catch ex As Threading.AbandonedMutexException
                hasHandle = True
            End Try

            If hasHandle = False Then
                log.Error(Set_ErrMSG(0, MSG_COM005))
                MsgBox(MSG_COM005, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, My.Application.Info.Title)
                Return
            End If

            'Oracle�ڑ�
            If OraConnect() = False Then
                log.Error(Set_ErrMSG(0, MSG_COM004))
                MsgBox(MSG_COM004, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, My.Application.Info.Title)
                Return
            End If

            '���喼�擾
            If My.Settings.���Ə��R�[�h <> 0 Then
                If DLTP0900_PROC0002("Sub_Main", gintSQLCODE, gstrSQLERRM) = False Then
                    MsgBox(gintSQLCODE & "-" & gstrSQLERRM & vbCr & MSG_COM902 & vbCr & MSG_COM901, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                    log.Error(Set_ErrMSG(gintSQLCODE, gstrSQLERRM))
                    Application.Exit()
                    Return
                End If
            End If

            '���Ə��ꗗ�擾
            If DLTP0901_PROC0001(gintSQLCODE, gstrSQLERRM) = False Then
                MsgBox(gintSQLCODE & "-" & gstrSQLERRM & vbCr & MSG_COM902 & vbCr & MSG_COM901, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                log.Error(Set_ErrMSG(gintSQLCODE, gstrSQLERRM))
                Application.Exit()
                Return
            End If

            '�A�v���N��
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
    ' *  ���W���[���@�\�F �w��t�H�[��ID�̃L���v�V������ݒ�
    ' *
    ' *  ���ӁA��������  �F�Ȃ�
    ' *  �����@�@�@�@�@�@�FFormID �E�E�e�t�H�[�����iDefine��`�j
    ' *  �@�@�@�@�@�@�@�@�FSenderName �E�E�e�t�H�[���I�u�W�F�N�g��
    ' *�@�ߒl�@�@�@�@�@�@�F�֐��E�E�t�H�[���L���v�V����
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Set_FormTitle(ByVal ProgramID As String,
                                         ByVal FormID As String) As String

        Try
            '�p�X�Ƀt�@�C������ǉ�
            Set_FormTitle = ProgramID & " " & FormID & " �y" & My.Application.Info.CompanyName & "�z"

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
#Region "�G���[���b�Z�[�W���` Set_ErrMSG"
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�G���[���b�Z�[�W���`
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@strErrCode    -- �G���[�R�[�h
    ' *�@�����Q�@�@�@�@�@�F�@strErrMessage -- �G���[���b�Z�[�W
    ' *�@�ߒl�@�@�@�@�@�@�F�@���`��G���[���b�Z�[�W
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Set_ErrMSG(ByVal strErrCode As Integer, ByVal strErrMessage As String) As String

        Dim strBuff As String

        strBuff = CType(strErrCode, String) & " " & strErrMessage

        Set_ErrMSG = strBuff

    End Function
#End Region

End Class
