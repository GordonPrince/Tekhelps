Imports System
Imports System.Windows.Forms
Imports AddinExpress.MSO
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

'Add-in Express Outlook Item Events Class
Public Class OutlookItemEventsClass1
    Inherits AddinExpress.MSO.ADXOutlookItemEvents

    Public Sub New(ByVal ADXModule As AddinExpress.MSO.ADXAddinModule)
        MyBase.New(ADXModule)
    End Sub

    Public Overrides Sub ProcessAttachmentAdd(ByVal Attachment As Object)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessAttachmentRead(ByVal Attachment As Object)
        ' MsgBox("ProcessAttachmentRead fired.")
    End Sub

    Public Overrides Sub ProcessBeforeAttachmentSave(ByVal Attachment As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessBeforeCheckNames(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessClose(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessCustomAction(ByVal Action As Object, ByVal Response As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessCustomPropertyChange(ByVal Name As String)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessForward(ByVal Forward As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessOpen(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessPropertyChange(ByVal Name As String)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessRead()
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessReply(ByVal Response As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessReplyAll(ByVal Response As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessSend(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessWrite(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessBeforeDelete(ByVal Item As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessAttachmentRemove(ByVal ByValAttachment As Object)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessBeforeAttachmentAdd(ByVal Attachment As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessBeforeAttachmentPreview(ByVal Attachment As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessBeforeAttachmentRead(ByVal attachment As Object, ByVal e As AddinExpress.MSO.ADXCancelEventArgs)
        Dim myAttachment As Microsoft.Office.Interop.Outlook.Attachment
        myAttachment = attachment
        If Left(myAttachment.DisplayName, 12) = "InstantFile_" Then
            MsgBox("This will open " & myAttachment.DisplayName & " instead of displaying the Note.")
            e.Cancel = True
        End If
    End Sub

    Public Overrides Sub ProcessBeforeAttachmentWriteToTempFile(ByVal Attachment As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessUnload()
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessBeforeAutoSave(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessBeforeRead()
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessAfterWrite()
        ' TODO: Add some code
    End Sub

End Class

