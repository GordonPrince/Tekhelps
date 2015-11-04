Imports System
Imports System.Windows.Forms
Imports AddinExpress.MSO
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

'Add-in Express Outlook Item Events Class
Public Class OutlookItemEventsClass1
    Inherits AddinExpress.MSO.ADXOutlookItemEvents

    Const strIFmatNo As String = "InstantFile_MatNo_"
    Const strIFdocNo As String = "InstantFile_DocNo_"
    Const strDocScanned As String = "Document scanned + imported:"
    Const strLastScanned As String = "LAST REQUESTED DOCUMENT scanned + imported:"

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
        If TypeOf Forward Is Outlook.MailItem Then
            Dim myMailItem As Outlook.MailItem = Forward
            ' Debug.Print("ProcessForward() myMailItem.BillingInformation = " & myMailItem.BillingInformation)
            myMailItem.BillingInformation = vbNullString
            Dim myAttachment As Outlook.Attachment
            For Each myAttachment In myMailItem.Attachments
                ' 11/3/2015 changed from VBA: If TypeOf myAttachment Is Outlook.Application And myAttachment.Class = 5 Then
                If Left(myAttachment.DisplayName, Len(strIFmatNo)) = strIFmatNo Then
                    If EmailMatNo(myAttachment, myMailItem.Subject) > 0 Then
                        Dim myUserProp As Outlook.UserProperty = myMailItem.UserProperties.Add("CameFromOutlook", Outlook.OlUserPropertyType.olText)
                        myUserProp.Value = "Forward"
                        Exit Sub
                    End If
                End If
            Next myAttachment
        End If
    End Sub

    Public Overrides Sub ProcessOpen(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessPropertyChange(ByVal Name As String)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessRead()
        ' MsgBox("ProcessRead fired")
    End Sub

    Public Overrides Sub ProcessReply(ByVal Response As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ReplyOrReplyAll(Response, "Reply")
    End Sub

    Public Overrides Sub ProcessReplyAll(ByVal Response As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ReplyOrReplyAll(Response, "ReplyAll")
    End Sub

    Private Sub ReplyOrReplyAll(Response As Object, strEventName As String)
        ' adds Outlook attachments from original message to Reply or ReplyAll
        Const strMsg As String = ".msg"
        Dim outlookApp As Outlook.Application, myResponse As Outlook.MailItem = Nothing
        Dim myInsp As Outlook.Inspector, myOriginal As Outlook.MailItem = Nothing
        Dim myAttachment As Outlook.Attachment, strFileName As String
        ' Dim myNoteA As Outlook.NoteItem
        Dim myUserProp As Outlook.UserProperty

        If TypeOf Response Is Outlook.MailItem Then
            myResponse = Response
            outlookApp = myResponse.Application
            If outlookApp.Inspectors.Count = 0 Then
                ' the user hit Reply from the Explorer window -- there's not item open in an Inspector window
                myOriginal = outlookApp.ActiveExplorer.Selection.Item(1)
                GoTo HaveItem
            Else
                For Each myInsp In outlookApp.Inspectors
                    myOriginal = myInsp.CurrentItem
                    If TypeOf myOriginal Is Outlook.MailItem Then
                        GoTo HaveItem
                    End If
                Next
            End If
            If myOriginal Is Nothing Then
                MsgBox("myOriginal is nothing.", vbEmpty, "ReplyOrReplyAll()")
                Exit Sub
            End If
HaveItem:
            'Debug.Print("myOriginal.Subject = " & myOriginal.Subject & ", myResponse.Subject = " & myResponse.Subject)
            'Debug.Print(myOriginal.Subject & " has " & myOriginal.Attachments.Count & " attachments.")

            ' 11/4/2015 Replying to an email that was Forwarded to you won't work without this
            Dim str1 As String = RemoveREFW(myOriginal.Subject)
            Dim str2 As String = RemoveREFW(myResponse.Subject)
            Debug.Print("str1 = " & str1)
            Debug.Print("str2 = " & str2)
            ' the first Reply puts "RE: " at the beginning of the new Subject, second Reply doesn't

            If InStr(str1, str2) > 0 Then
                For Each myAttachment In myOriginal.Attachments
                    If Right(LCase(myAttachment.FileName), 4) = strMsg Then
                        strFileName = "C:\tmp\" & myAttachment.FileName
                        myAttachment.SaveAsFile(strFileName)
                        myResponse.Attachments.Add(strFileName)
                        My.Computer.FileSystem.DeleteFile(strFileName)
                    End If
                Next myAttachment
                ' this is not in the Access code -- it's used to keep track of whether or not the email originated in InstantFile or Outlook
                myUserProp = myResponse.UserProperties.Add("CameFromOutlook", Outlook.OlUserPropertyType.olText)
                myUserProp.Value = strEventName
            End If
        End If
    End Sub

    Public Function RemoveREFW(strI As String) As String
        Do While Mid(strI, 3, 2) = ": " And Len(strI) > 4
            strI = Mid(strI, 5)
        Loop
        Return strI
    End Function

    Public Overrides Sub ProcessSend(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' ItemSend event of the Outlook.Application object
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
        Dim myAttachment As Outlook.Attachment
        Dim appAccess As Access.Application
        'Const strFile As String = "C:\tmp\myAttachment.msg"
        'Dim myNote As Outlook.NoteItem, strID As String

        myAttachment = attachment
        'Const strNewCallTrackingTag As String = "NewCall Tracking Item"
        'Const strNewCallAppointmentTag As String = "NewCall Appointment"
        'Const strIFtaskTag As String = "InstantFile_Task"

        Debug.Print(myAttachment.DisplayName)
        MsgBox("ProcessBeforeAttachmentRead fired.")
        If Left(myAttachment.DisplayName, Len(strIFdocNo)) = strIFdocNo Then
            Dim lngDocNo As Long = Mid(myAttachment.DisplayName, 19)
            If IsDBNull(lngDocNo) Or lngDocNo = 0 Then
                MsgBox("The item does not have a DocNo.", vbExclamation, "Show Document")
            Else
                appAccess = CType(Marshal.GetActiveObject("Access.Application"), Access.Application)
                If Not appAccess.Visible Then appAccess.Visible = True
                appAccess.Run("DisplayDocument", lngDocNo)
                e.Cancel = True
            End If
        ElseIf Left(myAttachment.DisplayName, Len(strIFmatNo)) = strIFmatNo Then
            Dim dblMatNo As Double = Mid(myAttachment.DisplayName, 19)
            If IsDBNull(dblMatNo) Or dblMatNo = 0 Then
                MsgBox("The item does not have a MatterNo.", vbExclamation, "Show Matter")
            Else
                appAccess = CType(Marshal.GetActiveObject("Access.Application"), Access.Application)
                If Not appAccess.Visible Then appAccess.Visible = True
                appAccess.Run("DisplayMatter", dblMatNo)
                e.Cancel = True
            End If
            ' these don't seem to work here -- copied to AdxOutlookAppEvents1_NewInspector
            'ElseIf Left(myAttachment.DisplayName, Len(strNewCallTrackingTag)) = strNewCallTrackingTag Then
            '    If My.Computer.FileSystem.FileExists(strFile) Then My.Computer.FileSystem.DeleteFile(strFile)
            '    myAttachment.SaveAsFile(strFile)
            '    myNote = myAttachment.Application.CreateItemFromTemplate(strFile)
            '    strID = Mid(myNote.Body, Len(strNewCallTrackingTag) + 3)
            '    ' Debug.Print(strID)
            '    If OpenItemFromID(myAttachment.Application, strID) Then
            '        e.Cancel = True
            '    End If
            'ElseIf Left(myAttachment.DisplayName, Len(strNewCallAppointmentTag)) = strNewCallAppointmentTag Then
            '    If My.Computer.FileSystem.FileExists(strFile) Then My.Computer.FileSystem.DeleteFile(strFile)
            '    myAttachment.SaveAsFile(strFile)
            '    myNote = myAttachment.Application.CreateItemFromTemplate(strFile)
            '    strID = Mid(myNote.Body, Len(strNewCallAppointmentTag) + 3)
            '    If OpenItemFromID(myAttachment.Application, strID) Then
            '        e.Cancel = True
            '    End If
            'ElseIf Left(myAttachment.DisplayName, Len(strIFtaskTag)) = strIFtaskTag Then
            '    If My.Computer.FileSystem.FileExists(strFile) Then My.Computer.FileSystem.DeleteFile(strFile)
            '    myAttachment.SaveAsFile(strFile)
            '    myNote = myAttachment.Application.CreateItemFromTemplate(strFile)
            '    strID = Mid(myNote.Body, Len(strIFtaskTag) + 3)
            '    If OpenItemFromID(myAttachment.Application, strID) Then
            '        e.Cancel = True
            '    End If
        End If
    End Sub

    Public Function OpenItemFromID(OutlookApp As Outlook.Application, strID As String) As Boolean
        Const strPublicFolders As String = "Public Folders"
        Dim olPublicFolder As Outlook.Folder, strPublicStoreID As String
        For Each olPublicFolder In OutlookApp.Session.Folders
            If Left(olPublicFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
                strPublicStoreID = olPublicFolder.StoreID
                For Each olFolder In olPublicFolder.Folders
                    If olFolder.Name = "All Public Folders" Then
                        Dim olNameSpace As Outlook.NameSpace = OutlookApp.GetNamespace("MAPI")
                        Dim item As Object = olNameSpace.GetItemFromID(strID, strPublicStoreID)
                        item.Display()
                        Return True
                    End If
                Next
            End If
        Next
        Return False
    End Function

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

    Private Function EmailMatNo(ByRef myAttach As Outlook.Attachment, ByVal strSubject As String) As Double
        On Error GoTo EmailMatNo_Error
        Dim strDisplayName As String
        Dim intX As Integer
        If Left(myAttach.DisplayName, 18) = strIFmatNo Then
            strDisplayName = Mid(myAttach.DisplayName, 19)
            intX = InStr(1, strDisplayName, Space(1))
            If intX > 0 Then strDisplayName = Left(strDisplayName, intX - 1)
            EmailMatNo = strDisplayName
        ElseIf Left(myAttach.DisplayName, 18) = strIFdocNo Then
            EmailMatNo = MatNoFromSubject(strSubject)
        Else
            EmailMatNo = False
        End If
        Exit Function

EmailMatNo_Error:
        MsgBox(Err.Description, vbExclamation, "Parse MatterNo from Attachment")
    End Function

    Private Function MatNoFromSubject(ByVal strSubject) As Double
        ' try to parse the MatterNo from the Subject line, not the attachment
        Dim intA As Integer, intB As Integer
        Dim strSearchFor As String = Nothing

        ' check for either string in the Subject. Use whichever one is found (changed 3/20/2006)
        intA = InStr(1, strSubject, strDocScanned)
        If intA > 0 Then
            strSearchFor = strDocScanned
        Else
            intA = InStr(1, strSubject, strLastScanned)
            If intA > 0 Then strSearchFor = strLastScanned
        End If
        If intA > 0 Then
            strSubject = Trim(Mid(strSubject, intA + Len(strSearchFor) + 1))
            intB = InStr(1, strSubject, Space(1))
            If intB > 0 Then
                On Error Resume Next
                MatNoFromSubject = Left(strSubject, intB)
                If Err.Number <> 0 Then
                    Err.Clear()
                    MatNoFromSubject = 0
                    On Error GoTo 0
                End If
            End If
        End If
    End Function
End Class


