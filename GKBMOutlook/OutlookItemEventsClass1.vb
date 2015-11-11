Imports System
Imports System.Windows.Forms
Imports AddinExpress.MSO
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

'Add-in Express Outlook Item Events Class
Public Class OutlookItemEventsClass1
    Inherits AddinExpress.MSO.ADXOutlookItemEvents

    Dim OutlookApp As Outlook.Application = CType(AddinModule.CurrentInstance, AddinModule).OutlookApp

    Public Sub New(ByVal ADXModule As AddinExpress.MSO.ADXAddinModule)
        MyBase.New(ADXModule)
    End Sub

    Public Overrides Sub ProcessAttachmentAdd(ByVal Attachment As Object)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessAttachmentRead(ByVal Attachment As Object)
        'Debug.Print("ProcessAttachmentRead fired")
    End Sub

    Public Overrides Sub ProcessBeforeAttachmentSave(ByVal Attachment As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        'Debug.Print("ProcessBeforeAttachmentSave fired")
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
            myMailItem.BillingInformation = vbNullString
            Dim myAttachments As Outlook.Attachments = myMailItem.Attachments
            Dim myAttachment As Outlook.Attachment = Nothing, x As Int16
            For x = 1 To myAttachments.Count
                myAttachment = myAttachments(x)
                If Left(myAttachment.DisplayName, Len(strIFmatNo)) = strIFmatNo Then
                    If EmailMatNo(myAttachment, myMailItem.Subject) > 0 Then
                        Dim myProps As Outlook.UserProperties = myMailItem.UserProperties
                        Dim myUserProp As Outlook.UserProperty = myProps.Add("CameFromOutlook", Outlook.OlUserPropertyType.olText)
                        myUserProp.Value = "Forward"
                        Try
                        Catch ex As System.Exception
                        Finally
                            Marshal.ReleaseComObject(myUserProp)
                            myUserProp = Nothing
                            Marshal.ReleaseComObject(myProps)
                            myProps = Nothing
                        End Try
                        Exit For
                    End If
                End If
            Next
            Try
            Catch ex As System.Exception
            Finally
                Marshal.ReleaseComObject(myAttachment)
                myAttachment = Nothing
                Marshal.ReleaseComObject(myAttachments)
                myAttachments = Nothing
                Marshal.ReleaseComObject(myMailItem)
                myMailItem = Nothing
            End Try
        End If
    End Sub

    Public Overrides Sub ProcessOpen(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' Debug.Print("ProcessOpen fired. Considering canceling display of the Note item here.")
    End Sub

    Public Overrides Sub ProcessPropertyChange(ByVal Name As String)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessRead()
        ' Debug.Print("ProcessRead fired")
    End Sub

    Public Overrides Sub ProcessReply(ByVal Response As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ReplyOrReplyAll(Response, "Reply")
    End Sub

    Public Overrides Sub ProcessReplyAll(ByVal Response As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ReplyOrReplyAll(Response, "ReplyAll")
    End Sub

    Private Sub ReplyOrReplyAll(Response As Object, strEventName As String)
        ' adds Outlook attachments from original message to Reply or ReplyAll
        If TypeOf Response Is Outlook.MailItem Then
        Else
            Return
        End If

        Dim myOriginal As Outlook.MailItem = Nothing
        Dim myResponse As Outlook.MailItem = Response
        Dim myInsps As Outlook.Inspectors = OutlookApp.Inspectors
        If myInsps.Count = 0 Then
            ' the user hit Reply from the Explorer window -- there's not an item open in an Inspector window
            ' 11/11/2015 changed myOriginal = OutlookApp.ActiveExplorer.Selection.Item(1)
            Dim myExpl As Outlook.Explorer = OutlookApp.ActiveExplorer
            Dim mySel As Outlook.Selection = myExpl.Selection
            myOriginal = mySel.Item(1)
            Marshal.ReleaseComObject(mySel)
            mySel = Nothing
            Marshal.ReleaseComObject(myExpl)
            myExpl = Nothing
        Else
            ' For Each myInsp In OutlookApp.Inspectors
            Dim myInsp As Outlook.Inspector = myInsps(1)
            myOriginal = myInsp.CurrentItem
            If TypeOf myOriginal Is Outlook.MailItem Then
            Else
                Marshal.ReleaseComObject(myInsp)
                myInsp = Nothing
                Marshal.ReleaseComObject(myOriginal)
                myOriginal = Nothing
                Marshal.ReleaseComObject(myInsps)
                myInsps = Nothing
                Marshal.ReleaseComObject(myResponse)
                myResponse = Nothing
                MsgBox("The item in the first Outlook Inspector window is not a MailItem.", vbOKOnly + vbExclamation, "ReplyOrReplyAll()")
                Return
            End If
            Marshal.ReleaseComObject(myInsp)
            myInsp = Nothing
        End If
        Marshal.ReleaseComObject(myInsps)
        myInsps = Nothing

        ' 11/4/2015 Replying to an email that was Forwarded to you won't work without this
        Dim str1 As String = RemoveREFW(myOriginal.Subject)
        Dim str2 As String = RemoveREFW(myResponse.Subject)

        ' the first Reply puts "RE: " at the beginning of the new Subject, second Reply doesn't
        If InStr(str1, str2) > 0 Then
            ' For Each myAttachment In myOriginal.Attachments
            Dim myAttachs As Outlook.Attachments = myOriginal.Attachments
            Dim x As Int16, myAttachment As Outlook.Attachment
            For x = 1 To myAttachs.Count
                myAttachment = myAttachs(x)
                If Right(LCase(myAttachment.FileName), 4) = ".msg" Then
                    Dim strFileName As String = "C:\tmp\" & myAttachment.FileName
                    myAttachment.SaveAsFile(strFileName)
                    Dim myRAs As Outlook.Attachments = myResponse.Attachments
                    myRAs.Add(strFileName)
                    Marshal.ReleaseComObject(myRAs)
                    myRAs = Nothing
                    Try
                        My.Computer.FileSystem.DeleteFile(strFileName)
                    Catch
                    End Try
                End If
            Next
            ' this is not in the Access code -- it's used to keep track of whether or not the email originated in InstantFile or Outlook
            Dim myProps As Outlook.UserProperties = myResponse.UserProperties
            Dim myUserProp As Outlook.UserProperty
            myUserProp = myProps.Add("CameFromOutlook", Outlook.OlUserPropertyType.olText)
            myUserProp.Value = strEventName
            Try
            Catch ex As Exception
            Finally
                Marshal.ReleaseComObject(myUserProp)
                myUserProp = Nothing
                Marshal.ReleaseComObject(myProps)
                myProps = Nothing
            End Try
        End If
        Try
        Catch ex As Exception
        Finally
            Marshal.ReleaseComObject(myOriginal)
            myOriginal = Nothing
            Marshal.ReleaseComObject(myResponse)
            myResponse = Nothing
        End Try
    End Sub

    Public Function RemoveREFW(strI As String) As String
        Do While Mid(strI, 3, 2) = ": " And Len(strI) > 4
            strI = Mid(strI, 5)
        Loop
        Return strI
    End Function

    Public Overrides Sub ProcessSend(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' ItemSend event of the object
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
        Debug.Print("ProcessBeforeAttachmentPreview")
    End Sub

    Public Overrides Sub ProcessBeforeAttachmentRead(ByVal attachment As Object, ByVal e As AddinExpress.MSO.ADXCancelEventArgs)
        Const strMsg As String = "This will only work if InstantFile is open." & vbNewLine & vbNewLine & _
                                 "Open InstantFile, then try this again."
        Dim myAttachment As Outlook.Attachment
        Dim appAccess As Access.Application = Nothing
        myAttachment = attachment
        'Debug.Print("ProcessBeforeAttachmentRead() " & Now & " TypeName(myAttachment) = " & TypeName(myAttachment))
        'Debug.Print("myAttachment.Type = " & myAttachment.Type)
        If Left(myAttachment.DisplayName, Len(strIFdocNo)) = strIFdocNo Then
            Const strDoc As String = "Open InstantFile Document"
            Dim lngDocNo As Long = Mid(myAttachment.DisplayName, 19)
            If IsDBNull(lngDocNo) Or lngDocNo = 0 Then
                MsgBox("The item does not have a DocNo.", vbExclamation, strDoc)
            Else
                Try
                    appAccess = CType(Marshal.GetActiveObject("Access.Application"), Access.Application)
                    If Not appAccess.Visible Then appAccess.Visible = True
                    appAccess.Run("DisplayDocument", lngDocNo)
                Catch
                    MsgBox(strMsg, vbExclamation + vbOKOnly, strDoc)
                Finally
                    Marshal.ReleaseComObject(appAccess)
                    appAccess = Nothing
                End Try
                e.Cancel = True
                Return
            End If
        ElseIf Left(myAttachment.DisplayName, Len(strIFmatNo)) = strIFmatNo Then
            Const strMat As String = "Show Matter in InstantFile"
            Dim dblMatNo As Double = Mid(myAttachment.DisplayName, 19)
            If IsDBNull(dblMatNo) Or dblMatNo = 0 Then
                MsgBox("The item does not have a MatterNo.", vbExclamation, strMat)
            Else
                Try
                    appAccess = CType(Marshal.GetActiveObject("Access.Application"), Access.Application)
                    If Not appAccess.Visible Then appAccess.Visible = True
                    appAccess.Run("DisplayMatter", dblMatNo)
                Catch
                    MsgBox(strMsg, vbExclamation + vbOKOnly, strMat)
                Finally
                    Marshal.ReleaseComObject(appAccess)
                    appAccess = Nothing
                End Try
                e.Cancel = True
                Return
            End If
        End If
    End Sub

    'Private Sub NAR(ByVal o As Object)
    '    ' copied from https://support.microsoft.com/en-us/kb/317109
    '    Try
    '        While (Marshal.ReleaseComObject(o) > 0)
    '        End While
    '    Catch
    '    Finally
    '        o = Nothing
    '    End Try
    'End Sub

    Public Overrides Sub ProcessBeforeAttachmentWriteToTempFile(ByVal Attachment As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        'Debug.Print("ProcessBeforeAttachmentPreview fired.")
    End Sub

    Public Overrides Sub ProcessUnload()
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessBeforeAutoSave(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' TODO: Add some code
    End Sub

    Public Overrides Sub ProcessBeforeRead()
        'Debug.Print("ProcessBeforeRead fired")
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


