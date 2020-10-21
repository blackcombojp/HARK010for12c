'/*-----------------------------------------------------------------------------
' * COPYRIGHT(C) ITI CORPORATION 2020
' * ITI CONFIDENTIAL AND PROPRIETARY
' *
' * All rights reserved by ITI Corporation.
' *-----------------------------------------------------------------------------/
Option Compare Binary
Option Explicit On
Option Strict On

Imports HARK010.HARK_Sub
Imports NAppUpdate.Framework
Public Class HARK_Update

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Shared Sub Update_CheckAfterRestart()

        For Each task As Tasks.IUpdateTask In UpdateManager.Instance.Tasks

            log.Info("TaskDescription:" & task.Description)
            log.Info("TaskExecutionStatus:" & task.ExecutionStatus.ToString)

        Next
    End Sub

    Public Shared Sub Update_Check()

        Try

            UpdateManager.Instance.BeginCheckForUpdates(New AsyncCallback(AddressOf Update_Check_Callback), Nothing)
            log.Info("BeginCheckForUpdates")

        Catch ex As Exception

            If TypeOf ex Is NAppUpdateException Then

                log.Error(Set_ErrMSG(Err.Number, ex.ToString))
                MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

                Return
            End If

        End Try

    End Sub

    Public Shared Sub Update_Check_Callback(ar As IAsyncResult)

        UpdateManager.Instance.EndCheckForUpdates(ar)

        log.Info("UpdatesAvailable is " & UpdateManager.Instance.UpdatesAvailable)

        If 0 = UpdateManager.Instance.UpdatesAvailable Then
            log.Info("This Exe is Latest edition")
            Return

        End If

        '更新
        Update_Prepare()

    End Sub

    Public Shared Sub Update_Prepare()

        gintMsg = MsgBox(MSG_UPD001, CType(MsgBoxStyle.OkOnly + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)

        UpdateManager.Instance.BeginPrepareUpdates(New AsyncCallback(AddressOf Update_Prepare_Callback), Nothing)
        log.Info("BeginPrepareUpdates")

    End Sub

    Public Shared Sub Update_Prepare_Callback(ar As IAsyncResult)

        Try

            UpdateManager.Instance.EndPrepareUpdates(ar)
            log.Info("EndPrepareUpdates")

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_UPD002 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

            Return
        End Try

        Update_Install()

    End Sub

    Public Shared Sub Update_Install()

        Try

            UpdateManager.Instance.ApplyUpdates(True, True, False)
            log.Info("ApplyUpdates")

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Sub

    Public Shared Sub Update_Rollback()

        Try

            UpdateManager.Instance.RollbackUpdates()
            log.Info("RollbackUpdates")

        Catch ex As Exception

            log.Error(Set_ErrMSG(Err.Number, ex.ToString))
            MsgBox(MSG_COM902 & vbCr & Err.Number & vbCr & ex.Message, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, My.Application.Info.Title)

        End Try

    End Sub

End Class
