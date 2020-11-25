'/*-----------------------------------------------------------------------------
' * COPYRIGHT(C) ITI CORPORATION 2020
' * ITI CONFIDENTIAL AND PROPRIETARY
' *
' * All rights reserved by ITI Corporation.
' *-----------------------------------------------------------------------------/
Option Compare Binary
Option Explicit On
Option Strict On

Imports Oracle.DataAccess.Client
'Imports Oracle.DataAccess.Types
Imports HARK010.HARK_Common
Imports HARK010.HARK_Sub

Public Class HARK_DBCommon

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Shared Oracomm As New OracleConnection
    Public Shared Oracmd As New OracleCommand
    Public Shared OraDr As OracleDataReader
    Public Shared OraDar As OracleDataAdapter
    Public Shared OraTran As OracleTransaction

    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@OraConnect
    ' *�@�N���X���@�@�@�@�F�@HASS_DBCommon
    ' *�@���W���[���@�\�@�F�@Oracle�ڑ�����
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@�Ȃ�
    ' *�@�ߒl�@�@�@�@�@�@�F�@True�E�E�����Afalse�E�E���s
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2006.3.5
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function OraConnect() As Boolean

        Dim StrParam As String

        Try

            OraConnect = False

            StrParam = "User id=" & My.Settings.DB���[�U & ";" & "Password=" & My.Settings.DB�p�X���[�h & ";" & "Data Source=" & My.Settings.DB�ڑ�������

            Oracomm.ConnectionString = StrParam

            Oracomm.Open()

            OraConnect = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@OraDisConnect
    ' *�@�N���X���@�@�@�@�F�@HASS_DBCommon
    ' *�@���W���[���@�\�@�F�@Oracle�ؒf����
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@�Ȃ�
    ' *�@�ߒl�@�@�@�@�@�@�F�@True�E�E�����Afalse�E�E���s
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2006.3.5
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function OraDisConnect() As Boolean

        Try

            OraDisConnect = False

            Oracomm.Close()

            OraDisConnect = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()
            Oracomm.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@OraConnectState
    ' *�@�N���X���@�@�@�@�F�@HASS_DBCommon
    ' *�@���W���[���@�\�@�F�@Oracle�ڑ���Ԋm�F
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@OraSessionID   -- �Z�b�V����ID�i�ߒl�j
    ' *�@�����Q�@�@�@�@�@�F�@SQLCODE        -- Oracle�G���[�R�[�h�i�ߒl�j
    ' *�@�����R�@�@�@�@�@�F�@SQLERRM        -- Oracle�G���[���b�Z�[�W�i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@True�E�E�ڑ����Afalse�E�E�ؒf
    ' *-----------------------------------------------------------------------------/
    Public Shared Function OraConnectState(ByRef PO_intOraSessionID As Integer,
                                           ByRef PO_intSQLCODE As Integer,
                                           ByRef PO_strSQLERRM As String) As Boolean

        Try

            Dim PO_01 As OracleParameter
            Dim PO_02 As OracleParameter
            Dim PO_03 As OracleParameter
            Dim PO_04 As OracleParameter

            OraConnectState = False
            PO_intOraSessionID = 0

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0998S.PROC0001"
            Oracmd.CommandType = CommandType.StoredProcedure

            'Output�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.Varchar2, 255, DBNull.Value, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Int32, ParameterDirection.Output)
            PO_04 = Oracmd.Parameters.Add("PO_04", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��Call
            Oracmd.ExecuteNonQuery()

            PO_intSQLCODE = CInt(PO_03.Value.ToString)
            PO_strSQLERRM = PO_04.Value.ToString

            '���^�[���R�[�h�ł̏����U�蕪��
            If PO_intSQLCODE = 0 Then

                PO_intOraSessionID = CInt(PO_02.Value.ToString)
                OraConnectState = True

            Else

                log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

            End If

            Exit Function

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))

            Return OraConnectState

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@���Ə��ꗗ�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@Program_ID   -- �v���O����_ID
    ' *�@�����Q�@�@�@�@�@�F�@SQLCODE      -- Oracle�G���[�R�[�h�i�ߒl�j
    ' *�@�����R�@�@�@�@�@�F�@SQLERRM      -- Oracle�G���[���b�Z�[�W�i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@True -- ����擾 False -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DLTP0901_PROC0001(ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Dim i As Integer

        Try
            DLTP0901_PROC0001 = False

            gint���Ə�Cnt = 0

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0901.PROC0001"
            Oracmd.CommandType = CommandType.StoredProcedure

            'Output�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��Call
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            ���Ə�Array = Nothing
            gint���Ə�Cnt = 0

            '���^�[���R�[�h�ł̏����U�蕪��
            If PO_intSQLCODE = 0 Then

                i = 0

                While OraDr.Read

                    '�������Ď擾
                    ReDim Preserve ���Ə�Array(i)

                    '�O���[�o���ϐ��ɃZ�b�g
                    ���Ə�Array(i).int���Ə��R�[�h = OraDr.GetInt32(0)
                    ���Ə�Array(i).str���Ə��� = OraDr.GetString(1)
                    i += 1

                End While

                gint���Ə�Cnt = i

            Else

                log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                Exit Function

            End If

            DLTP0901_PROC0001 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@���喼�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@Program_ID   -- �v���O����_ID
    ' *�@�����Q�@�@�@�@�@�F�@SQLCODE      -- Oracle�G���[�R�[�h�i�ߒl�j
    ' *�@�����R�@�@�@�@�@�F�@SQLERRM      -- Oracle�G���[���b�Z�[�W�i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@True -- ����擾 False -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DLTP0900_PROC0002(ByVal PI_strProgram_ID As String,
                                             ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim PI_01 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Try
            DLTP0900_PROC0002 = False

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0900.PROC0002"
            Oracmd.CommandType = CommandType.StoredProcedure

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = My.Settings.���Ə��R�[�h

            'Output�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.Varchar2, 60, DBNull.Value, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��Call
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            gstr���喼 = Nothing

            '���^�[���R�[�h�ł̏����U�蕪��
            If PO_intSQLCODE = 0 Then

                gstr���喼 = PO_01.Value.ToString

            Else

                log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                Exit Function

            End If

            DLTP0900_PROC0002 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�T�u�v���O�����ꗗ�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@SQLCODE      -- Oracle�G���[�R�[�h�i�ߒl�j
    ' *�@�����Q�@�@�@�@�@�F�@SQLERRM      -- Oracle�G���[���b�Z�[�W�i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@True -- ����擾 False -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0000_PROC0001(ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim PI_01 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Dim i As Integer

        Try
            DTNP0000_PROC0001 = False

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DTNP0000.PROC0001"
            Oracmd.CommandType = CommandType.StoredProcedure

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = My.Settings.���Ə��R�[�h

            'Output�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��Call
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            �T�u�v���O����Array = Nothing
            gint�T�u�v���O����Cnt = 0

            '���^�[���R�[�h�ł̏����U�蕪��
            If PO_intSQLCODE = 0 Then

                i = 0

                While OraDr.Read

                    '�������Ď擾
                    ReDim Preserve �T�u�v���O����Array(i)

                    '�O���[�o���ϐ��ɃZ�b�g
                    �T�u�v���O����Array(i).str�T�u�v���O�����R�[�h = OraDr.GetString(0)
                    �T�u�v���O����Array(i).str�T�u�v���O������ = OraDr.GetString(1)
                    i += 1

                End While

                gint�T�u�v���O����Cnt = i

            Else

                log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                Exit Function

            End If

            DTNP0000_PROC0001 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F  �v���O�������擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@Program_ID   -- �v���O����_ID
    ' *�@�����Q�@�@�@�@�@�F�@SQLCODE      -- Oracle�G���[�R�[�h�i�ߒl�j
    ' *�@�����R�@�@�@�@�@�F�@SQLERRM      -- Oracle�G���[���b�Z�[�W�i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@True -- ����擾 False -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0000_PROC0002(ByVal PI_strProgram_ID As String,
                                             ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Try
            DTNP0000_PROC0002 = False

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DTNP0000.PROC0002"
            Oracmd.CommandType = CommandType.StoredProcedure

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Varchar2, ParameterDirection.Input)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = My.Settings.���Ə��R�[�h
            PI_02.Value = PI_strProgram_ID

            'Output�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��Call
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            gudt�v���O�����}�X�^.IsClear()

            '���^�[���R�[�h�ł̏����U�蕪��
            If PO_intSQLCODE = 0 Then

                OraDr.Read()

                If OraDr.IsDBNull(0) = False Then gudt�v���O�����}�X�^.str�����֐� = OraDr.GetString(0)
                If OraDr.IsDBNull(1) = False Then gudt�v���O�����}�X�^.str�o�̓w�b�_ = OraDr.GetString(1)
                If OraDr.IsDBNull(2) = False Then gudt�v���O�����}�X�^.str�o�͋�ؕ��� = OraDr.GetString(2)
                If OraDr.IsDBNull(3) = False Then gudt�v���O�����}�X�^.int���������P = OraDr.GetInt32(3)
                If OraDr.IsDBNull(4) = False Then gudt�v���O�����}�X�^.int���������Q = OraDr.GetInt32(4)
                If OraDr.IsDBNull(5) = False Then gudt�v���O�����}�X�^.int���������R = OraDr.GetInt32(5)
                If OraDr.IsDBNull(6) = False Then gudt�v���O�����}�X�^.int���������S = OraDr.GetInt32(6)
                If OraDr.IsDBNull(7) = False Then gudt�v���O�����}�X�^.int���������T = OraDr.GetInt32(7)
                If OraDr.IsDBNull(8) = False Then gudt�v���O�����}�X�^.int���������U = OraDr.GetInt32(8)
                If OraDr.IsDBNull(9) = False Then gudt�v���O�����}�X�^.int���������V = OraDr.GetInt32(9)
                If OraDr.IsDBNull(10) = False Then gudt�v���O�����}�X�^.int���������W = OraDr.GetInt32(10)
                If OraDr.IsDBNull(11) = False Then gudt�v���O�����}�X�^.int���������X = OraDr.GetInt32(11)
                If OraDr.IsDBNull(12) = False Then gudt�v���O�����}�X�^.int���������P�O = OraDr.GetInt32(12)
                If OraDr.IsDBNull(13) = False Then gudt�v���O�����}�X�^.str���������P�q���g = OraDr.GetString(13)
                If OraDr.IsDBNull(14) = False Then gudt�v���O�����}�X�^.str���������Q�q���g = OraDr.GetString(14)
                If OraDr.IsDBNull(15) = False Then gudt�v���O�����}�X�^.str���������R�q���g = OraDr.GetString(15)
                If OraDr.IsDBNull(16) = False Then gudt�v���O�����}�X�^.str���������S�q���g = OraDr.GetString(16)
                If OraDr.IsDBNull(17) = False Then gudt�v���O�����}�X�^.str���������T�q���g = OraDr.GetString(17)
                If OraDr.IsDBNull(18) = False Then gudt�v���O�����}�X�^.str���������U�q���g = OraDr.GetString(18)
                If OraDr.IsDBNull(19) = False Then gudt�v���O�����}�X�^.str���������V�q���g = OraDr.GetString(19)
                If OraDr.IsDBNull(20) = False Then gudt�v���O�����}�X�^.str���������W�q���g = OraDr.GetString(20)
                If OraDr.IsDBNull(21) = False Then gudt�v���O�����}�X�^.str���������X�q���g = OraDr.GetString(21)
                If OraDr.IsDBNull(22) = False Then gudt�v���O�����}�X�^.str���������P�O�q���g = OraDr.GetString(22)
                If OraDr.IsDBNull(23) = False Then gudt�v���O�����}�X�^.int�T�u�v���O����_ID = OraDr.GetInt32(23)

            Else

                log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                Exit Function

            End If

            DTNP0000_PROC0002 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@���̎����ꗗ�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@Program_ID   -- �v���O����_ID
    ' *�@�����Q�@�@�@�@�@�F�@�敪         -- �������ʋ敪
    ' *�@�����R�@�@�@�@�@�F�@SQLCODE      -- Oracle�G���[�R�[�h�i�ߒl�j
    ' *�@�����S�@�@�@�@�@�F�@SQLERRM      -- Oracle�G���[���b�Z�[�W�i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@True -- ����擾 False -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0000_PROC0003(ByVal PI_strProgram_ID As String,
                                             ByVal PI_�敪 As Integer,
                                             ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Dim i As Integer

        Try
            DTNP0000_PROC0003 = False

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DTNP0000.PROC0003"
            Oracmd.CommandType = CommandType.StoredProcedure

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Varchar2, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Int32, ParameterDirection.Input)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = My.Settings.���Ə��R�[�h
            PI_02.Value = PI_strProgram_ID
            PI_03.Value = PI_�敪

            'Output�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��Call
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            Select Case PI_�敪

                Case 1

                    ���Ӑ�Array = Nothing
                    gint���Ӑ�Cnt = 0

                    '���^�[���R�[�h�ł̏����U�蕪��
                    If PO_intSQLCODE = 0 Then

                        i = 0

                        While OraDr.Read

                            '�������Ď擾
                            ReDim Preserve ���Ӑ�Array(i)

                            '�O���[�o���ϐ��ɃZ�b�g
                            ���Ӑ�Array(i).lng���Ӑ�R�[�h = OraDr.GetInt64(0)
                            ���Ӑ�Array(i).str���Ӑ於 = OraDr.GetString(1)
                            i += 1

                        End While

                        gint���Ӑ�Cnt = i

                    Else

                        log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                        Exit Function

                    End If

                Case 2

                    ���v��Array = Nothing
                    gint���v��Cnt = 0

                    '���^�[���R�[�h�ł̏����U�蕪��
                    If PO_intSQLCODE = 0 Then

                        i = 0

                        While OraDr.Read

                            '�������Ď擾
                            ReDim Preserve ���v��Array(i)

                            '�O���[�o���ϐ��ɃZ�b�g
                            ���v��Array(i).lng���v��R�[�h = OraDr.GetInt64(0)
                            ���v��Array(i).str���v�於 = OraDr.GetString(1)
                            i += 1

                        End While

                        gint���v��Cnt = i

                    Else

                        log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                        Exit Function

                    End If

            End Select

            DTNP0000_PROC0003 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�o�׌��i�������f�[�^����
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@strProgram_ID     -- �v���O����_ID
    ' *  �����Q�@�@�@�@�@�F�@�T�u�v���O����_ID -- �T�u�v���O����_ID
    ' *  �����R�@�@�@�@�@�F�@���Ӑ�R�[�h      -- ���Ӑ�R�[�h
    ' *  �����S�@�@�@�@�@�F�@���v��R�[�h      -- ���v��R�[�h
    ' *�@�����T�@�@�@�@�@�F�@Dgv               -- DataGridView�i�ߒl�j
    ' *�@�����U�@�@�@�@�@�F�@ROWCount          -- �����i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@0 -- ����擾 2 -- ���R�[�h�� 9 -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC001(ByVal PI_strProgram_ID As String,
                                            ByVal PI_int�T�u�v���O����_ID As Integer,
                                            ByVal PI_lng���Ӑ�R�[�h As Long,
                                            ByVal PI_lng���v��R�[�h As Long,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PI_04 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC001 = False

            OraDs.Clear()

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            OraTran = Oracomm.BeginTransaction

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int64, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Int64, ParameterDirection.Input)
            PI_04 = Oracmd.Parameters.Add("PI_04", OracleDbType.Int32, ParameterDirection.Input)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = PI_int�T�u�v���O����_ID
            PI_02.Value = PI_lng���Ӑ�R�[�h
            If PI_lng���v��R�[�h = 0 Then
                PI_03.Value = vbNullString
            Else
                PI_03.Value = PI_lng���v��R�[�h
            End If
            PI_04.Value = My.Settings.���Ə��R�[�h

            '�A�E�g�v�b�g�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��call
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            OraTran.Commit()

            DTNP0403_PROC001 = True

        Catch Oraex As OracleException

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            If Not IsNothing(OraTran) Then
                OraTran.Dispose()
            End If

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@Oliver�G���[�f�[�^����
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@strProgram_ID     -- �v���O����_ID
    ' *  �����Q�@�@�@�@�@�F�@�Ώۓ�            -- �����Ώۓ�
    ' *�@�����R�@�@�@�@�@�F�@Dgv               -- DataGridView�i�ߒl�j
    ' *�@�����S�@�@�@�@�@�F�@ROWCount          -- �����i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@0 -- ����擾 2 -- ���R�[�h�� 9 -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC003(ByVal PI_strProgram_ID As String,
                                            ByVal PI_�Ώۓ� As String,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC003 = False

            OraDs.Clear()

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            OraTran = Oracomm.BeginTransaction

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = PI_�Ώۓ�.Trim

            '�A�E�g�v�b�g�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��call
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            OraTran.Commit()

            DTNP0403_PROC003 = True

        Catch Oraex As OracleException

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            If Not IsNothing(OraTran) Then
                OraTran.Dispose()
            End If

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�����ݏo�ԍ���񌟍�
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@strProgram_ID     -- �v���O����_ID
    ' *  �����Q�@�@�@�@�@�F�@���Ӑ�R�[�h      -- ���Ӑ�R�[�h
    ' *  �����R�@�@�@�@�@�F�@���v��R�[�h      -- ���v��R�[�h
    ' *  �����S�@�@�@�@�@�F�@���i�R�[�h        -- ���i�R�[�h
    ' *  �����T�@�@�@�@�@�F�@�����i��        -- �����i��
    ' *�@�����U�@�@�@�@�@�F�@Dgv               -- DataGridView�i�ߒl�j
    ' *�@�����V�@�@�@�@�@�F�@ROWCount          -- �����i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@0 -- ����擾 2 -- ���R�[�h�� 9 -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC004(ByVal PI_strProgram_ID As String,
                                            ByVal PI_lng���Ӑ�R�[�h As Long,
                                            ByVal PI_lng���v��R�[�h As Long,
                                            ByVal PI_str���i�R�[�h As String,
                                            ByVal PI_str�����i�� As String,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PI_04 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC004 = False

            OraDs.Clear()

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int64, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int64, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Varchar2, 60, DBNull.Value, ParameterDirection.Input)
            PI_04 = Oracmd.Parameters.Add("PI_04", OracleDbType.Varchar2, 60, DBNull.Value, ParameterDirection.Input)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = PI_lng���Ӑ�R�[�h
            PI_02.Value = PI_lng���v��R�[�h
            If IsNull(PI_str���i�R�[�h) Then
                PI_03.Value = vbNullString
            Else
                PI_03.Value = PI_str���i�R�[�h
            End If
            If IsNull(PI_str�����i��) Then
                PI_04.Value = vbNullString
            Else
                PI_04.Value = PI_str�����i��
            End If

            '�A�E�g�v�b�g�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��call
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            DTNP0403_PROC004 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@PHsmos�󒍏�񌟍�
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@strProgram_ID     -- �v���O����_ID
    ' *  �����Q�@�@�@�@�@�F�@�����R�[�h        -- �����R�[�h
    ' *  �����Q�@�@�@�@�@�F�@���Ӑ�R�[�h      -- ���Ӑ�R�[�h
    ' *  �����R�@�@�@�@�@�F�@�Ώۓ�            -- �Ώۓ�
    ' *�@�����S�@�@�@�@�@�F�@Dgv               -- DataGridView�i�ߒl�j
    ' *�@�����T�@�@�@�@�@�F�@ROWCount          -- �����i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@0 -- ����擾 2 -- ���R�[�h�� 9 -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC006(ByVal PI_strProgram_ID As String,
                                            ByVal PI_int�����R�[�h As Integer,
                                            ByVal PI_lng���Ӑ�R�[�h As Long,
                                            ByVal PI_�Ώۓ� As String,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC006 = False

            OraDs.Clear()

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int64, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = PI_int�����R�[�h
            PI_02.Value = PI_lng���Ӑ�R�[�h
            If PI_�Ώۓ�.Trim.CompareTo(DUMMY_DATESTRING) = 0 Then
                PI_03.Value = vbNullString
            Else
                PI_03.Value = PI_�Ώۓ�.Trim
            End If

            '�A�E�g�v�b�g�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��call
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            DTNP0403_PROC006 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�Z�b�V�������폜
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@strProgram_ID   -- �v���O����_ID
    ' *�@�����Q�@�@�@�@�@�F�@strV�敪        -- SUB�Z�b�V�������.V�敪
    ' *�@�����R�@�@�@�@�@�F�@intN�敪        -- SUB�Z�b�V�������.N�敪
    ' *�@�ߒl�@�@�@�@�@�@�F�@0 -- ����I�� 9 -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DLTP0998S_PROC0013(ByVal strProgram_ID As String,
                                              ByVal strV�敪 As String,
                                              ByVal intN�敪 As Integer) As Integer


        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PI_04 As OracleParameter
        Dim PI_05 As OracleParameter
        Dim PI_06 As OracleParameter
        Dim PI_07 As OracleParameter
        Dim PI_08 As OracleParameter
        Dim PI_09 As OracleParameter
        Dim PI_10 As OracleParameter
        Dim PI_11 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_99 As OracleParameter

        Try

            DLTP0998S_PROC0013 = 9

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0998S.PROC0013"
            Oracmd.CommandType = CommandType.StoredProcedure

            'OraTran = Oracomm.BeginTransaction

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)
            PI_04 = Oracmd.Parameters.Add("PI_04", OracleDbType.Int32, ParameterDirection.Input)
            PI_05 = Oracmd.Parameters.Add("PI_05", OracleDbType.Int32, ParameterDirection.Input)
            PI_06 = Oracmd.Parameters.Add("PI_06", OracleDbType.Int32, ParameterDirection.Input)
            PI_07 = Oracmd.Parameters.Add("PI_07", OracleDbType.Int32, ParameterDirection.Input)
            PI_08 = Oracmd.Parameters.Add("PI_08", OracleDbType.Int32, ParameterDirection.Input)
            PI_09 = Oracmd.Parameters.Add("PI_09", OracleDbType.Int32, ParameterDirection.Input)
            PI_10 = Oracmd.Parameters.Add("PI_10", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)
            PI_11 = Oracmd.Parameters.Add("PI_11", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)

            'Output�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.Varchar2, 255, DBNull.Value, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_99 = Oracmd.Parameters.Add("PO_99", OracleDbType.Int32, ParameterDirection.Output)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = "DELETE"
            PI_02.Value = strProgram_ID
            PI_03.Value = strV�敪
            PI_04.Value = intN�敪
            PI_05.Value = vbNullString
            PI_06.Value = vbNullString
            PI_07.Value = vbNullString
            PI_08.Value = vbNullString
            PI_09.Value = vbNullString
            PI_10.Value = vbNullString
            PI_11.Value = vbNullString

            '�X�g�A�h�v���V�[�W��call
            Oracmd.ExecuteNonQuery()

            '���^�[���R�[�h�ł̏����U�蕪��
            If CInt(PO_99.Value.ToString) = 0 Then

                DLTP0998S_PROC0013 = 0

                Exit Function

            End If

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@���v��ʔ�����o��
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@strProgram_ID     -- �v���O����_ID
    ' *  �����Q�@�@�@�@�@�F�@�����R�[�h        -- �����R�[�h
    ' *  �����R�@�@�@�@�@�F�@���Ӑ�R�[�h      -- ���Ӑ�R�[�h
    ' *�@�����S�@�@�@�@�@�F�@Dgv               -- DataGridView�i�ߒl�j
    ' *�@�����T�@�@�@�@�@�F�@ROWCount          -- �����i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@0 -- ����擾 2 -- ���R�[�h�� 9 -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC007(ByVal PI_strProgram_ID As String,
                                            ByVal PI_int�����R�[�h As Integer,
                                            ByVal PI_lng���Ӑ�R�[�h As Long,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC007 = False

            OraDs.Clear()

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int64, ParameterDirection.Input)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = PI_int�����R�[�h
            PI_02.Value = PI_lng���Ӑ�R�[�h

            '�A�E�g�v�b�g�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��call
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            DTNP0403_PROC007 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�L�������ؔ���񌟍�
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@strProgram_ID     -- �v���O����_ID
    ' *  �����Q�@�@�@�@�@�F�@���Ӑ�R�[�h      -- ���Ӑ�R�[�h
    ' *  �����R�@�@�@�@�@�F�@�Ώۓ�            -- �Ώۓ�
    ' *�@�����S�@�@�@�@�@�F�@Dgv               -- DataGridView�i�ߒl�j
    ' *�@�����T�@�@�@�@�@�F�@ROWCount          -- �����i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@0 -- ����擾 2 -- ���R�[�h�� 9 -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DTNP0403_PROC008(ByVal PI_strProgram_ID As String,
                                            ByVal PI_lng���Ӑ�R�[�h As Long,
                                            ByVal PI_�Ώۓ� As String,
                                            ByRef PO_Dgv As DataGridView,
                                            ByRef PO_intROWCount As Integer) As Boolean

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim OraDs As New DataSet()

        Try
            DTNP0403_PROC008 = False

            OraDs.Clear()

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = PI_strProgram_ID
            Oracmd.CommandType = CommandType.StoredProcedure

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int64, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Varchar2, 20, DBNull.Value, ParameterDirection.Input)

            '�C���v�b�g�l�ݒ�
            PI_01.Value = PI_lng���Ӑ�R�[�h
            PI_02.Value = PI_�Ώۓ�.Trim

            '�A�E�g�v�b�g�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��call
            OraDar = New OracleDataAdapter(Oracmd)
            OraDar.Fill(OraDs, "TMP")
            '����`�F�b�N�{�b�N�X�ǉ�
            OraDs.Tables("TMP").Columns.Add("���", Type.GetType("System.Boolean")).DefaultValue = False

            PO_Dgv.DataSource = OraDs.Tables("TMP")
            PO_intROWCount = PO_Dgv.RowCount

            DTNP0403_PROC008 = True

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F  ���[�Ǘ����擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@Program_ID         -- �v���O����_ID
    ' *�@�����Q�@�@�@�@�@�F�@SPD�V�X�e���R�[�h  -- SPD�V�X�e���R�[�h
    ' *�@�����R�@�@�@�@�@�F�@�T�u�v���O����_ID  -- �T�u�v���O����_ID
    ' *�@�����S�@�@�@�@�@�F�@SQLCODE            -- Oracle�G���[�R�[�h�i�ߒl�j
    ' *�@�����T�@�@�@�@�@�F�@SQLERRM            -- Oracle�G���[���b�Z�[�W�i�ߒl�j
    ' *�@�ߒl�@�@�@�@�@�@�F�@0 -- ����擾 1 -- �f�[�^���� 9 -- �G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function DLTP0996S_PROC0001(ByVal PI_strProgram_ID As String,
                                             ByVal PI_intSPD�V�X�e���R�[�h As Integer,
                                             ByVal PI_int�T�u�v���O����_ID As Integer,
                                             ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Integer

        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PI_03 As OracleParameter
        Dim PI_04 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter
        Dim PO_03 As OracleParameter

        Try
            DLTP0996S_PROC0001 = 9

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0996S.PROC0001"
            Oracmd.CommandType = CommandType.StoredProcedure

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Varchar2, ParameterDirection.Input)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int32, ParameterDirection.Input)
            PI_03 = Oracmd.Parameters.Add("PI_03", OracleDbType.Int32, ParameterDirection.Input)
            PI_04 = Oracmd.Parameters.Add("PI_04", OracleDbType.Int32, ParameterDirection.Input)


            '�C���v�b�g�l�ݒ�
            PI_01.Value = PI_strProgram_ID
            PI_02.Value = PI_intSPD�V�X�e���R�[�h
            PI_03.Value = PI_int�T�u�v���O����_ID
            PI_04.Value = My.Settings.���Ə��R�[�h

            'Output�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.RefCursor, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Int32, ParameterDirection.Output)
            PO_03 = Oracmd.Parameters.Add("PO_03", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            '�X�g�A�h�v���V�[�W��Call
            OraDr = Oracmd.ExecuteReader()

            PO_intSQLCODE = CInt(PO_02.Value.ToString)
            PO_strSQLERRM = PO_03.Value.ToString

            gudt���[�Ǘ����.IsClear()

            '���^�[���R�[�h�ł̏����U�蕪��
            Select Case PO_intSQLCODE

                Case 0

                    OraDr.Read()

                    If OraDr.IsDBNull(0) = False Then gudt���[�Ǘ����.lng���[�Ǘ��ԍ� = OraDr.GetInt64(0)
                    If OraDr.IsDBNull(1) = False Then gudt���[�Ǘ����.str���[�� = OraDr.GetString(1)
                    If OraDr.IsDBNull(2) = False Then gudt���[�Ǘ����.str�e���v���[�g�� = OraDr.GetString(2)
                    If OraDr.IsDBNull(3) = False Then gudt���[�Ǘ����.str�����֐� = OraDr.GetString(3)
                    If OraDr.IsDBNull(4) = False Then gudt���[�Ǘ����.int�v���r���[�t���O = OraDr.GetInt32(4)
                    If OraDr.IsDBNull(5) = False Then gudt���[�Ǘ����.int�o�͌`���敪 = OraDr.GetInt32(5)
                    If OraDr.IsDBNull(6) = False Then gudt���[�Ǘ����.str�V�[�g���P = OraDr.GetString(6)
                    If OraDr.IsDBNull(7) = False Then gudt���[�Ǘ����.int�ő喾�׍s���P = OraDr.GetInt32(7)
                    If OraDr.IsDBNull(8) = False Then gudt���[�Ǘ����.int���׊Ԋu�s���P = OraDr.GetInt32(8)
                    If OraDr.IsDBNull(9) = False Then gudt���[�Ǘ����.str�V�[�g���Q = OraDr.GetString(9)
                    If OraDr.IsDBNull(10) = False Then gudt���[�Ǘ����.int�ő喾�׍s���Q = OraDr.GetInt32(10)
                    If OraDr.IsDBNull(11) = False Then gudt���[�Ǘ����.int���׊Ԋu�s���Q = OraDr.GetInt32(11)
                    If OraDr.IsDBNull(12) = False Then gudt���[�Ǘ����.str�V�[�g���R = OraDr.GetString(12)
                    If OraDr.IsDBNull(13) = False Then gudt���[�Ǘ����.int�ő喾�׍s���R = OraDr.GetInt32(13)
                    If OraDr.IsDBNull(14) = False Then gudt���[�Ǘ����.int���׊Ԋu�s���R = OraDr.GetInt32(14)
                    If OraDr.IsDBNull(15) = False Then gudt���[�Ǘ����.str�V�[�g���S = OraDr.GetString(15)
                    If OraDr.IsDBNull(16) = False Then gudt���[�Ǘ����.int�ő喾�׍s���S = OraDr.GetInt32(16)
                    If OraDr.IsDBNull(17) = False Then gudt���[�Ǘ����.int���׊Ԋu�s���S = OraDr.GetInt32(17)
                    If OraDr.IsDBNull(18) = False Then gudt���[�Ǘ����.int�o�[�R�[�h��� = OraDr.GetInt32(18)
                    If OraDr.IsDBNull(19) = False Then gudt���[�Ǘ����.int�o�[�R�[�h���� = OraDr.GetInt32(19)
                    If OraDr.IsDBNull(20) = False Then gudt���[�Ǘ����.int�o�[�R�[�h�� = OraDr.GetInt32(20)
                    If OraDr.IsDBNull(21) = False Then gudt���[�Ǘ����.int�\���{�� = OraDr.GetInt32(21)
                    If OraDr.IsDBNull(22) = False Then gudt���[�Ǘ����.str�T�v = OraDr.GetString(22)
                    If OraDr.IsDBNull(23) = False Then gudt���[�Ǘ����.str���l = OraDr.GetString(23)

                    DLTP0996S_PROC0001 = PO_intSQLCODE

                    Exit Function

                Case 1

                    DLTP0996S_PROC0001 = PO_intSQLCODE

                    Exit Function

                Case Else

                    log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

                    Exit Function

            End Select

        Catch Oraex As OracleException

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            Oracmd.Dispose()

        End Try

    End Function

    ''' <summary>
    ''' �L�������ؔ��f�[�^�o�͌��ʂ����Ƀe�[�u�����X�V
    ''' </summary>
    ''' <param name="rdocDgv">����</param>
    ''' <param name="PO_intSQLCODE">Oracle�G���[�R�[�h�i�ߒl�j</param>
    ''' <param name="PO_strSQLERRM">Oracle�G���[���b�Z�[�W�i�ߒl�j</param>
    ''' <returns>True -- ����I�� False -- �G���[</returns>
    Public Shared Function DLTP0201_PROC0024(ByVal rdocDgv As DataGridView,
                                             ByRef PO_intSQLCODE As Integer,
                                             ByRef PO_strSQLERRM As String) As Boolean

        Dim i As Integer
        Dim intCnt As Integer
        Dim intPrintChk As Integer
        Dim ID() As Integer = Nothing
        Dim PI_01 As OracleParameter
        Dim PI_02 As OracleParameter
        Dim PO_01 As OracleParameter
        Dim PO_02 As OracleParameter

        Try

            DLTP0201_PROC0024 = False

            '�X�g�A�h�v���V�[�W���ݒ�
            Oracmd = Oracomm.CreateCommand()
            Oracmd.CommandText = "DLTP0201.PROC0024"
            Oracmd.CommandType = CommandType.StoredProcedure

            intCnt = 0

            For i = 0 To rdocDgv.RowCount - 1

                If CInt(rdocDgv.Rows(i).Cells(21).Value) = 9 Then
                    Continue For
                End If

                '����`�F�b�N�{�b�N�X
                If IsDBNull(rdocDgv.Rows(i).Cells(24).Value) = True Then
                    intPrintChk = 0
                Else
                    intPrintChk = CInt(rdocDgv.Rows(i).Cells(24).Value)
                End If

                '����`�F�b�N�{�b�N�XOFF�͑ΏۊO
                If intPrintChk = 0 Then
                    Continue For
                End If

                ReDim Preserve ID(intCnt)
                ID(intCnt) = CInt(rdocDgv.Rows(i).Cells(23).Value)
                intCnt += 1
            Next

            If intCnt = 0 Then Return True

            OraTran = Oracomm.BeginTransaction

            '�C���v�b�g�p�����[�^�ݒ�
            PI_01 = Oracmd.Parameters.Add("PI_01", OracleDbType.Int32)
            PI_02 = Oracmd.Parameters.Add("PI_02", OracleDbType.Int32, ParameterDirection.Input)

            'Output�p�����[�^�ݒ�
            PO_01 = Oracmd.Parameters.Add("PO_01", OracleDbType.Int32, ParameterDirection.Output)
            PO_02 = Oracmd.Parameters.Add("PO_02", OracleDbType.Varchar2, 1024, DBNull.Value, ParameterDirection.Output)

            PI_01.CollectionType = OracleCollectionType.PLSQLAssociativeArray
            PI_01.Size = intCnt

            '�C���v�b�g�l�ݒ�
            PI_01.Value = ID
            PI_02.Value = intCnt

            '�X�g�A�h�v���V�[�W��Call
            Oracmd.ExecuteNonQuery()

            PO_intSQLCODE = CType(PO_01.Value.ToString, Integer)
            PO_strSQLERRM = PO_02.Value.ToString

            '���^�[���R�[�h�ł̏����U�蕪��
            If PO_intSQLCODE = 0 Then

                OraTran.Commit()

                DLTP0201_PROC0024 = True

                Exit Function

            End If

            log.Error(Set_ErrMSG(PO_intSQLCODE, PO_strSQLERRM))

            OraTran.Rollback()

        Catch Oraex As OracleException

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Oraex.Number, Oraex.ToString))
            Throw

        Catch ex As Exception

            If Not IsNothing(OraTran) Then
                OraTran.Rollback()
            End If

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        Finally

            If Not IsNothing(OraTran) Then
                OraTran.Dispose()
            End If

            Oracmd.Dispose()

        End Try

    End Function

End Class
