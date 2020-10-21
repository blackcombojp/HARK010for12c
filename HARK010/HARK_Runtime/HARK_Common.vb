'/*-----------------------------------------------------------------------------
' * COPYRIGHT(C) ITI CORPORATION 2020
' * ITI CONFIDENTIAL AND PROPRIETARY
' *
' * All rights reserved by ITI Corporation.
' *-----------------------------------------------------------------------------/
Option Compare Binary
Option Explicit On
Option Strict On

'Imports System.IO.FileStream
Imports System.IO
Imports HARK010.HARK_Sub
Public Class HARK_Common

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Get_AppPath
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@��exe�p�X�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@�Ȃ�
    ' *�@�ߒl�@�@�@�@�@�@�F�@��exe�p�X
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2005.12.5
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_AppPath() As String

        Dim strFullAppName As String
        Dim strFullAppPath As String

        Try
            '��exe�p�X�擾(�t�@�C�����܂�)
            strFullAppName = Reflection.Assembly.GetExecutingAssembly().Location

            '��exe��������(�E�[[\]���폜)
            strFullAppPath = Path.GetDirectoryName(strFullAppName)

            Get_AppPath = strFullAppPath

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Get_AppFileName
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@��exe���i�g���q�t���j�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@�Ȃ�
    ' *�@�ߒl�@�@�@�@�@�@�F�@��exe���i�g���q�t���j
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2005.12.5
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_AppFileName() As String

        Dim strFullAppName As String
        Dim strFullAppPath As String

        Try
            '��exe�p�X�擾(�t�@�C�����܂�)
            strFullAppName = Reflection.Assembly.GetExecutingAssembly().Location

            '��exe�p�X������(�E�[[\]���폜)
            strFullAppPath = Path.GetFileName(strFullAppName)

            Get_AppFileName = strFullAppPath

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Chk_FileExists
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@�w��t�@�C�����݃`�F�b�N
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@FileName�E�E�w��t�@�C����(�t���p�X)
    ' *�@�ߒl�@�@�@�@�@�@�F�@Ture�E�E���݂���,False�E�E���݂��Ȃ�
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2005.3.20
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Chk_FileExists(ByVal FileName As String) As Boolean

        Try

            Chk_FileExists = False

            If File.Exists(FileName) = True Then

                Chk_FileExists = True

            End If

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Chk_DirExists
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@�w��t�H���_���݃`�F�b�N
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@FileName�E�E�w��t�H���_��(�t���p�X)
    ' *�@�ߒl�@�@�@�@�@�@�F�@Ture�E�E���݂���,False�E�E���݂��Ȃ�
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2007.8.23
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Chk_DirExists(ByVal DirName As String) As Boolean

        Try

            Chk_DirExists = False

            If Right(DirName.Trim, 1) <> "\" Then
                DirName &= "\"
            End If

            If Directory.Exists(DirName) = True Then

                Chk_DirExists = True

            End If

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Set_FilePath
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@�w��Dir�p�X�Ɏw��t�@�C������ǉ�
    ' *
    ' *�@���ӁA���������@�F�@DirPath�͉E�[[\]�͕s�v
    ' *�@�����P�@�@�@�@�@�F�@DirPath �E�E�C�ӂ�Dir�p�X
    ' *�@�����Q�@�@�@�@�@�F�@FileName �E�E�C�ӂ̃t�@�C����(�g���q�t)
    ' *�@�ߒl�@�@�@�@�@�@�F�@�t���p�X�t�@�C����
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2005.3.11
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Set_FilePath(ByVal DirPath As String, ByVal Filename As String) As String

        Try
            '�p�X�Ƀt�@�C������ǉ�
            Set_FilePath = Path.Combine(DirPath, Filename)

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Chk_Version
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@�w��exe�̃o�[�W������r
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@Path�E�E���sexe�̃t���p�X 
    ' *�@�ߒl�@�@�@�@�@�@�F�@True�E�E�o�[�W�����A�b�v�K�v�AFalse�E�E�s�v
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2005.12.5
    ' *�@�C�������@�@�@�@�F
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Chk_Version(ByVal Path As String) As Boolean

        Dim strAppVer As String         '���s�o�[�W����
        Dim strAppVerUp As String       '�����o�[�W����

        Try

            Chk_Version = False

            '�t�@�C�����݃`�F�b�N
            If Chk_FileExists(Path) = False Then

                Exit Function

            End If

            '���s�o�[�W�����擾
            strAppVer = Application.ProductVersion

            '�����o�[�W�����擾
            Dim vi As FileVersionInfo = FileVersionInfo.GetVersionInfo(Path)

            strAppVerUp = vi.FileVersion

            '�o�[�W������r
            If strAppVer >= strAppVerUp Then
                Exit Function
            End If

            Chk_Version = True

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@IsNull
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@������NULL�`�F�b�N
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@text�E�E�w�蕶���� 
    ' *�@�ߒl�@�@�@�@�@�@�F�@True�E�ENULL�AFalse�E�ENULL�łȂ�
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2005.12.5
    ' *�@�C�������@�@�@�@�F
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function IsNull(ByVal text As String) As Boolean

        If IsNothing(text.Trim) = True Then
            Return True
        End If

        If text.Trim Is Nothing = True Then
            Return True
        End If

        If text.Trim = "" Then
            Return True
        End If

        Return False

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@IsNumeric
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@������ː����ϊ��ۃ`�F�b�N
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@text�E�E�w�蕶���� 
    ' *�@�ߒl�@�@�@�@�@�@�F�@True�E�E�ϊ��AFalse�E�E�ϊ��s��
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2005.12.5
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function IsNumeric(ByVal text As String) As Boolean

        Try
            Double.Parse(text)
            Return True
        Catch
            Return False
        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@IsDate
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@������˓��t�ϊ��ۃ`�F�b�N
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@text�E�E�w�蕶���� 
    ' *�@�ߒl�@�@�@�@�@�@�F�@True�E�E�ϊ��AFalse�E�E�ϊ��s��
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2005.12.5
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function IsDate(ByVal text As String) As Boolean

        Try
            text = text.Replace("_", "")
            Date.Parse(text)
            Return True
        Catch
            Return False
        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Get_DesktopPath
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@�f�X�N�g�b�v�p�X�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@�Ȃ� 
    ' *�@�ߒl�@�@�@�@�@�@�F�@�f�X�N�g�b�v�p�X
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2005.12.5
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_DesktopPath() As String

        Dim Path As String

        Try

            Path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)

            Get_DesktopPath = Path

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Entry_Check
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@������̓��̓`�F�b�N
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@strData   ---- �`�F�b�N���镶����                                                     
    ' *�@�����Q�@�@�@�@�@�F�@SizeFlg   ---- 1 �S�p 2 ���p 0 ���p               
    ' *�@�����R�@�@�@�@�@�F�@StyleFlg  ---- 0 �����̂� 1 �p�����̂� 2 ���̑�   
    ' *�@�����S�@�@�@�@�@�F�@intLength ---- �ő啶����                     
    ' *�@�ߒl�@�@�@�@�@�@�F�@Boolean
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@S.Matsuo
    ' *�@�쐬���@�@�@�@�@�F�@2006.9.1
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Entry_Check(ByVal strData As String, ByVal intSizeFlg As Integer, ByVal intStyleFlg As Integer, ByVal intMaxLength As Integer) As Boolean

        Dim Base() As Byte
        Dim UniBase() As Byte
        Dim Code As Text.Encoding
        Dim intByte As Integer

        Try
            Entry_Check = False

            '��������`�F�b�N
            If intSizeFlg < 0 Or intSizeFlg > 2 Then
                Exit Function
            End If

            If intStyleFlg < 0 Or intStyleFlg > 2 Then
                Exit Function
            End If

            If intMaxLength = 0 Then
                Exit Function
            End If

            '�ϊ������R�[�h�w��
            Code = Text.Encoding.GetEncoding(932)

            '������̃o�C�g���擾
            Base = Text.Encoding.Unicode.GetBytes(strData)

            'Unicode�o�C�g�֕ϊ�
            UniBase = Text.Encoding.Convert(Text.Encoding.Unicode, Code, Base)

            '�擾������̃o�C�g�����
            intByte = CType(UniBase.Length, Integer)

            '�����񒷔�r
            If intByte > intMaxLength Then
                Exit Function
            End If

            If intSizeFlg = 1 Then        '�S�p�`�F�b�N
                If Text.RegularExpressions.Regex.IsMatch(strData, "^[a-zA-Z0-9�-�!-/]+") Then
                    Exit Function
                End If
            ElseIf intSizeFlg = 2 Then    '���p�`�F�b�N
                If Text.RegularExpressions.Regex.IsMatch(strData, "^[a-zA-Z0-9�-�!-/\uFF61-\uFF9F]+") = False Then
                    Exit Function
                End If
            End If

            If intStyleFlg = 0 Then    '�����̂�
                If Not Text.RegularExpressions.Regex.IsMatch(strData, "^[0-9]{1," & intMaxLength & "}$") Then
                    Exit Function
                End If
            ElseIf intStyleFlg = 1 Then     '�p�����̂�
                If Not Text.RegularExpressions.Regex.IsMatch(strData, "^[a-zA-Z0-9\-/]{1," & intMaxLength & "}$") Then
                    Exit Function
                End If
            ElseIf intStyleFlg = 2 Then     '���̑�
                If Not Text.RegularExpressions.Regex.IsMatch(strData, "^[a-zA-Z0-9�-�!-/\uFF61-\uFF9F]{1," & intMaxLength & "}$") Then
                    Exit Function
                End If
            End If

            Entry_Check = True

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@���t�̓��̓`�F�b�N
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@strDate   ---- �`�F�b�N������t                                                     
    ' *�@�����Q�@�@�@�@�@�F�@ChkFlg    ---- 1�E�EYYYY/MM/DD 2�E�EYYYY/MM                
    ' *�@�ߒl�@�@�@�@�@�@�F�@Boolean
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Date_Check(ByVal strDate As String, ByVal ChkFlg As Integer) As Boolean

        Dim strBuff() As String    '/�p�z��
        Dim iLastDay As Integer
        Dim iYear As Integer
        Dim iMonth As Integer
        Dim iDate As Integer

        Try

            Date_Check = False
            gstrDate = ""

            If Entry_Check(strDate, 2, 1, 10) = False Then
                Exit Function
            End If

            If Text.RegularExpressions.Regex.IsMatch(strDate, "^[0-9/]{1,10}$") = False Then
                Exit Function
            End If

            strBuff = Split(strDate, "/")

            If ChkFlg = 1 Then
                Select Case strBuff.Length

                    Case Is = 1

                        If strDate.Length <> 8 Then
                            Exit Function
                        End If

                        iYear = CType(strDate.Substring(0, 4), Integer)
                        iMonth = CType(strDate.Substring(4, 2), Integer)
                        iDate = CType(strDate.Substring(6, 2), Integer)

                        If Text.RegularExpressions.Regex.IsMatch(strDate, "^[12][0-9]{7}$") Then

                            If (Date.MinValue.Year > iYear) OrElse (iYear > Date.MaxValue.Year) Then
                                Exit Function
                            End If

                            If (Date.MinValue.Month > iMonth) OrElse (iMonth > Date.MaxValue.Month) Then
                                Exit Function
                            End If

                            iLastDay = Date.DaysInMonth(iYear, iMonth)

                            If (Date.MinValue.Day > iDate) OrElse (iDate > iLastDay) Then
                                Exit Function
                            End If

                            gstrDate = strDate.Substring(0, 4) & "/" &
                                       strDate.Substring(4, 2) & "/" &
                                       strDate.Substring(6, 2)

                            Date_Check = True
                            Exit Function

                        Else
                            Exit Function
                        End If

                    Case Is = 3

                        iYear = CType(strBuff(0), Integer)
                        iMonth = CType(strBuff(1), Integer)
                        iDate = CType(strBuff(2), Integer)

                        If Text.RegularExpressions.Regex.IsMatch(strDate, "^[12][0-9]{1,3}[/][0-9]{1,2}[/][0-9]{1,2}$") Then

                            If strBuff(1).Length = 1 And CDbl(strBuff(1)) < 10 Then
                                strBuff(1) = "0" & strBuff(1)
                            End If

                            If strBuff(2).Length = 1 And CDbl(strBuff(2)) < 10 Then
                                strBuff(2) = "0" & strBuff(2)
                            End If

                            If (Date.MinValue.Year > iYear) OrElse (iYear > Date.MaxValue.Year) Then
                                Exit Function
                            End If

                            If (Date.MinValue.Month > iMonth) OrElse (iMonth > Date.MaxValue.Month) Then
                                Exit Function
                            End If

                            iLastDay = Date.DaysInMonth(iYear, iMonth)

                            If (Date.MinValue.Day > iDate) OrElse (iDate > iLastDay) Then
                                Exit Function
                            End If

                            '����I��
                            gstrDate = strBuff(0) & "/" &
                                       strBuff(1) & "/" &
                                       strBuff(2)

                            Date_Check = True
                            Exit Function

                        Else
                            Exit Function
                        End If

                End Select

            ElseIf ChkFlg = 2 Then
                Select Case strBuff.Length

                    Case Is = 1

                        If strDate.Length <> 6 Then
                            Exit Function
                        End If

                        iYear = CType(strDate.Substring(0, 4), Integer)
                        iMonth = CType(strDate.Substring(4, 2), Integer)
                        ' iDate = CType(strDate.Substring(6, 2), Integer)

                        If Text.RegularExpressions.Regex.IsMatch(strDate, "^[12][0-9]{5}$") Then

                            If (Date.MinValue.Year > iYear) OrElse (iYear > Date.MaxValue.Year) Then
                                Exit Function
                            End If

                            If (Date.MinValue.Month > iMonth) OrElse (iMonth > Date.MaxValue.Month) Then
                                Exit Function
                            End If

                            gstrDate = strDate.Substring(0, 4) & "/" &
                                       strDate.Substring(4, 2)

                            Date_Check = True
                            Exit Function

                        Else
                            Exit Function
                        End If

                    Case Is = 2

                        iYear = CType(strBuff(0), Integer)
                        iMonth = CType(strBuff(1), Integer)
                        '         iDate = CType(strBuff(2), Integer)


                        If Text.RegularExpressions.Regex.IsMatch(strDate, "^[12][0-9]{1,3}[/][0-9]{1,2}$") Then

                            If strBuff(1).Length = 1 And CDbl(strBuff(1)) < 10 Then
                                strBuff(1) = "0" & strBuff(1)
                            End If

                            If (Date.MinValue.Year > iYear) OrElse (iYear > Date.MaxValue.Year) Then
                                Exit Function
                            End If

                            If (Date.MinValue.Month > iMonth) OrElse (iMonth > Date.MaxValue.Month) Then
                                Exit Function
                            End If


                            '����I��
                            gstrDate = strBuff(0) & "/" &
                                       strBuff(1)

                            Date_Check = True
                            Exit Function

                        Else
                            Exit Function
                        End If

                End Select

            End If

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Get_DirectoryName
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@�f�B���N�g�����擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@��΃p�X
    ' *�@�ߒl�@�@�@�@�@�@�F�@�f�B���N�g����
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2007.11.24
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_DirectoryName(ByVal strPath As String) As String

        Get_DirectoryName = Path.GetDirectoryName(strPath)

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Get_Extension
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@�g���q�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@��΃p�X
    ' *�@�ߒl�@�@�@�@�@�@�F�@�g���q
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2007.11.24
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_Extension(ByVal strPath As String) As String

        Get_Extension = Right(Path.GetExtension(strPath), 3)

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Get_ExtensionEx
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@�g���q�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@��΃p�X
    ' *�@�ߒl�@�@�@�@�@�@�F�@�g���q
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2019.06.21
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_ExtensionEx(ByVal strPath As String) As String

        Get_ExtensionEx = Right(Path.GetExtension(strPath), 4)

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Get_FileName
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@�t�@�C�����擾�i�g���q�t�j
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@��΃p�X
    ' *�@�ߒl�@�@�@�@�@�@�F�@�t�@�C�����i�g���q�t�j
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2007.11.24
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_FileName(ByVal strPath As String) As String

        Get_FileName = Path.GetFileName(strPath)

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Get_FileNameWithoutExtension
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@�t�@�C�����擾�i�g���q���j
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@��΃p�X
    ' *�@�ߒl�@�@�@�@�@�@�F�@�t�@�C�����i�g���q���j
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2007.11.24
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_FileNameWithoutExtension(ByVal strPath As String) As String

        Get_FileNameWithoutExtension = Path.GetFileNameWithoutExtension(strPath)

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Get_PathRoot
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@���[�g�f�B���N�g�����擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@��΃p�X
    ' *�@�ߒl�@�@�@�@�@�@�F�@���[�g�f�B���N�g����
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2007.11.24
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_PathRoot(ByVal strPath As String) As String

        Get_PathRoot = Path.GetPathRoot(strPath)

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[�����@�@�F�@Nvl
    ' *�@�N���X���@�@�@�@�F�@HASS_Common
    ' *�@���W���[���@�\�@�F�@NULL�u��
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@obj1�E�E�w��I�u�W�F�N�g
    ' *�@�����P�@�@�@�@�@�F�@obj2�E�E�w��I�u�W�F�N�g
    ' *�@�ߒl�@�@�@�@�@�@�F�@obj1��NULL�Ȃ�obj2�Aobj1��NOTNULL�Ȃ�obj1
    ' *
    ' *�@�쐬�ҁ@�@�@�@�@�F�@k.takada
    ' *�@�쐬���@�@�@�@�@�F�@2008.1.26
    ' *�@�C�������@�@�@�@�F�@
    ' *
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Nvl(ByVal obj1 As Object, ByVal obj2 As Object) As Object

        Try

            If obj1 Is Nothing Then
                Nvl = obj2
                Return Nvl
            Else
                Nvl = obj1
                Return Nvl
            End If


            If IsNothing(obj1) Then
                Nvl = obj2
                Return Nvl
            Else
                Nvl = obj1
                Return Nvl
            End If

            Return CType(vbNullString, Object)

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@NULL�u��(String�p)
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@obj1 -- �w��I�u�W�F�N�g
    ' *�@�����Q�@�@�@�@�@�F�@obj2 -- �w��I�u�W�F�N�g
    ' *�@�ߒl�@�@�@�@�@�@�F�@obj1��NULL�Ȃ�obj2 obj1��NOTNULL�Ȃ�obj1
    ' *-----------------------------------------------------------------------------/
    Public Shared Function NvlString(ByVal obj1 As String, ByVal obj2 As String) As String

        Try
            If IsNull(obj1) Then
                NvlString = obj2
                Return NvlString
            Else
                NvlString = obj1
                Return NvlString
            End If

            Return vbNullString

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�f�B���N�g���쐬
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@��΃p�X
    ' *�@�ߒl�@�@�@�@�@�@�F�@true -- OK false -- NG
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Create_Dir(ByVal strPath As String) As Boolean

        Dim hstream As DirectoryInfo

        Try
            hstream = Directory.CreateDirectory(strPath)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�w��f�B���N�g�����S�t�@�C�����폜
    ' *
    ' *�@���ӁA���������@�F�@
    ' *�@�����P�@�@�@�@�@�F�@DirName -- �C�ӂ̃f�B���N�g��
    ' *�@�ߒl�@�@�@�@�@�@�F�@0 -- ����I�� 1 -- �ُ�I��
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Delete_Files(ByVal DirName As String) As Integer

        Dim Files As String() = Directory.GetFiles(DirName)
        Dim f As String

        Try
            For Each f In Files

                Dim FileName As String = Set_FilePath(DirName, Path.GetFileName(f))

                Kill(FileName)

            Next f
            Return 0
        Catch ex As Exception
            Return 1
        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�t�@�C���e��`�F�b�N
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@��΃p�X
    ' *�@�ߒl�@�@�@�@�@�@�F�@true -- OK�Afalse -- NG
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Check_FilePath(ByVal strPath As String) As Boolean

        Dim tmpFileName As String
        Dim invalidChars As Char() = Path.GetInvalidPathChars


        Try
            '�f�B���N�g�����݃`�F�b�N
            If Not Directory.Exists(Get_DirectoryName(strPath)) Then
                Return False
            End If

            '�f�B���N�g���A�N�Z�X�`�F�b�N
            If Not Create_DummyFile(Set_FilePath(Get_DirectoryName(strPath), DUMMY_FILENAME)) Then
                Return False
            End If

            '�_�~�[�t�@�C���폜
            Kill(Set_FilePath(Get_DirectoryName(strPath), DUMMY_FILENAME))

            '�v���b�g�t�H�[���ŗL�����`�F�b�N
            tmpFileName = Get_FileNameWithoutExtension(strPath)
            If tmpFileName.IndexOfAny(invalidChars) >= 0 Then
                Return False
            End If

            '�g���q���݃`�F�b�N
            If Not Has_Extension(strPath) Then
                Return False
            End If

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�_�~�[�t�@�C������
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@��΃p�X
    ' *�@�ߒl�@�@�@�@�@�@�F�@true -- OK false -- NG
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Create_DummyFile(ByVal strPath As String) As Boolean

        Dim hstream As FileStream = Nothing

        Try
            Try
                hstream = File.Create(strPath)
                Return True
            Catch ex As Exception
                Return False
            Finally
                If Not hstream Is Nothing Then
                    hstream.Close()
                End If
            End Try
        Catch ex As Exception
            Return False
        Finally
            If Not hstream Is Nothing Then
                Dim hDisposable As IDisposable = hstream
                hDisposable.Dispose()
            End If
        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�g���q���݃`�F�b�N
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@��΃p�X
    ' *�@�ߒl�@�@�@�@�@�@�F�@true -- ���݂��� false -- ���݂��Ȃ�
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Has_Extension(ByVal strPath As String) As Boolean

        Has_Extension = Path.HasExtension(strPath)

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�w��t�@�C�����폜
    ' *
    ' *�@���ӁA���������@�F�@
    ' *�@�����P�@�@�@�@�@�F�@FileName -- �C�ӂ̃t�@�C����(�g���q�t)
    ' *�@�ߒl�@�@�@�@�@�@�F�@0 -- ����I�� 55 -- �t�@�C���I�[�v�� 53 -- �t�@�C�����݃G���[ 99 -- ���̑��G���[
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Delete_File(ByVal Filename As String) As Integer

        Try
            '�t�@�C���폜
            Kill(Filename)
            Return 0
        Catch ex As FileNotFoundException
            '�t�@�C�������݂��Ȃ�
            Return 53
        Catch ex As IOException
            '�t�@�C�������v���Z�X�Ŏg�p��
            Return 55
        Catch ex As Exception
            '���̑��G���[
            Return 99
        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�R���s���[�^���擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@�Ȃ�
    ' *�@�ߒl�@�@�@�@�@�@�F�@�R���s���[�^��
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_MachineName() As String

        Try

            Get_MachineName = Environment.MachineName

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�R���s���[�^���O�C�����[�U�[���擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@�Ȃ�
    ' *�@�ߒl�@�@�@�@�@�@�F�@�R���s���[�^���O�C�����[�U�[��
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_LoginUserName() As String

        Try

            Get_LoginUserName = Environment.UserName

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�R���s���[�^IP�A�h���X�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@�Ȃ�
    ' *�@�ߒl�@�@�@�@�@�@�F�@IP�A�h���X
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_IPAddress() As String

        Try
            'VISTA�ȍ~API�d�l�ύX�ׁ̈AOS�o�[�W�������̎擾
            Select Case Get_OSVersion()

                Case OS_WINDOWSVISTA, OS_WINDOWS7, OS_WINDOWS8

                    Get_IPAddress = Net.Dns.GetHostEntry(Net.Dns.GetHostName).AddressList(1).ToString
                    Exit Select

                Case OS_MACINTOSH, OS_UNIX, OS_UNKNOWN, OS_WINDOWS32s, OS_WINDOWSCE, OS_XBOX

                    Get_IPAddress = ""
                    Exit Select

                Case Else

                    Get_IPAddress = Net.Dns.GetHostEntry(Net.Dns.GetHostName).AddressList(0).ToString
                    Exit Select

            End Select

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@OS�o�[�W�����擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@�Ȃ�
    ' *�@�ߒl�@�@�@�@�@�@�F�@OS�o�[�W����
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_OSVersion() As Integer

        Dim os As OperatingSystem = Environment.OSVersion

        Try

            Select Case os.Platform

                Case PlatformID.Win32Windows

                    If os.Version.Major >= 4 Then

                        Select Case os.Version.Minor
                            Case 0
                                Get_OSVersion = OS_WINDOWS95
                                Exit Select
                            Case 10
                                Get_OSVersion = OS_WINDOWS98
                                Exit Select
                            Case 90
                                Get_OSVersion = OS_WINDOWSME
                                Exit Select
                        End Select

                    End If
                    Exit Select

                Case PlatformID.Win32NT

                    Select Case os.Version.Major
                        Case 3
                            Select Case os.Version.Minor
                                Case 0
                                    Get_OSVersion = OS_WINDOWSNT3
                                    Exit Select
                                Case 1
                                    Get_OSVersion = OS_WINDOWSNT31
                                    Exit Select
                                Case 5
                                    Get_OSVersion = OS_WINDOWSNT35
                                    Exit Select
                                Case 51
                                    Get_OSVersion = OS_WINDOWSNT351
                                    Exit Select
                            End Select
                            Exit Select
                        Case 4
                            If os.Version.Minor = 0 Then
                                Get_OSVersion = OS_WINDOWSNT4
                            End If
                            Exit Select
                        Case 5
                            Select Case os.Version.Minor
                                Case 0
                                    Get_OSVersion = OS_WINDOWS2000
                                    Exit Select
                                Case 1
                                    Get_OSVersion = OS_WINDOWSXP
                                    Exit Select
                                Case 2
                                    Get_OSVersion = OS_WINDOWSSERVER2003
                                    Exit Select
                            End Select
                            Exit Select
                        Case 6
                            Select Case os.Version.Minor
                                Case 0
                                    Get_OSVersion = OS_WINDOWSVISTA 'Windows Server 2008�܂�
                                    Exit Select
                                Case 1
                                    Get_OSVersion = OS_WINDOWS7  'Windows Server 2008 R2�܂�
                                    Exit Select
                                Case 2
                                    Get_OSVersion = OS_WINDOWS8  'Windows Server 2012�܂�
                                    Exit Select
                                Case 3
                                    Get_OSVersion = OS_WINDOWS81  'Windows Server 2012 R2�܂�
                                    Exit Select
                            End Select
                            Exit Select
                        Case 10
                            Select Case os.Version.Minor
                                Case 0
                                    Get_OSVersion = OS_WINDOWS10 'Windows Server 2016�܂�
                                    Exit Select
                            End Select
                            Exit Select
                    End Select
                    Exit Select

                Case PlatformID.Win32S
                    Get_OSVersion = OS_WINDOWS32s
                    Exit Select

                Case PlatformID.WinCE
                    Get_OSVersion = OS_WINDOWSCE
                    Exit Select

                Case PlatformID.Unix
                    '.NET Framework 2.0�ȍ~
                    Get_OSVersion = OS_UNIX
                    Exit Select

                Case PlatformID.Xbox
                    '.NET Framework 3.5�ȍ~
                    Get_OSVersion = OS_XBOX
                    Exit Select

                Case PlatformID.MacOSX
                    '.NET Framework 3.5�ȍ~
                    Get_OSVersion = OS_MACINTOSH
                    Exit Select

                Case Else
                    Get_OSVersion = OS_UNKNOWN
                    Exit Select

            End Select

            Return Get_OSVersion

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
    '/*-----------------------------------------------------------------------------
    ' *�@���W���[���@�\�@�F�@�J�����g���[�U�[ApplicationData�p�X�擾
    ' *
    ' *�@���ӁA���������@�F�@�Ȃ�
    ' *�@�����P�@�@�@�@�@�F�@�Ȃ� 
    ' *�@�ߒl�@�@�@�@�@�@�F�@�f�X�N�g�b�v�p�X
    ' *-----------------------------------------------------------------------------/
    Public Shared Function Get_ApplicationPath() As String

        Dim Path As String

        Try

            Path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

            Get_ApplicationPath = Path

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            Throw ex

        End Try

    End Function
End Class
