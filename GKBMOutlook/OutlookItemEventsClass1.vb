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
        Else
            Return
        End If

        Dim myMailItem As Outlook.MailItem = Nothing
        Dim myAttachments As Outlook.Attachments = Nothing
        Dim myAttachment As Outlook.Attachment = Nothing
        Dim myProps As Outlook.UserProperties = Nothing
        Dim myUserProp As Outlook.UserProperty = Nothing
        Try
            myMailItem = Forward
            myMailItem.BillingInformation = vbNullString
            myAttachments = myMailItem.Attachments
            Dim x As Short
            For x = 1 To myAttachments.Count
                myAttachment = myAttachments(x)
                If Left(myAttachment.DisplayName, Len(strIFmatNo)) = strIFmatNo Then
                    If EmailMatNo(myAttachment, myMailItem.Subject) > 0 Then
                        myProps = myMailItem.UserProperties
                        myUserProp = myProps.Add("CameFromOutlook", Outlook.OlUserPropertyType.olText)
                        myUserProp.Value = "Forward"
                        Marshal.ReleaseComObject(myUserProp)
                        Marshal.ReleaseComObject(myProps)
                        Return
                    End If
                End If
                Marshal.ReleaseComObject(myAttachment)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "Update forwarded E-mail with InstantFile tags")

        Finally
            If myUserProp IsNot Nothing Then Marshal.ReleaseComObject(myUserProp) : myUserProp = Nothing
            If myProps IsNot Nothing Then Marshal.ReleaseComObject(myProps) : myProps = Nothing
            If myAttachment IsNot Nothing Then Marshal.ReleaseComObject(myAttachment) : myAttachment = Nothing
            If myAttachments IsNot Nothing Then Marshal.ReleaseComObject(myAttachments) : myAttachments = Nothing
            ' Marshal.ReleaseComObject(myMailItem) : myMailItem = Nothing
        End Try
    End Sub

    Public Overrides Sub ProcessOpen(ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        ' Debug.Print("ProcessOpen() fired")
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

        Dim myResponse As Outlook.MailItem = Nothing
        Dim myOriginal As Outlook.MailItem = Nothing
        Dim myInsps As Outlook.Inspectors = Nothing
        Dim myExpl As Outlook.Explorer = Nothing
        Dim mySel As Outlook.Selection = Nothing
        Dim myInsp As Outlook.Inspector = Nothing
        Dim myProperties As Outlook.UserProperties = Nothing
        Dim myProp As Outlook.UserProperty = Nothing

        Dim str1 As String
        Dim str2 As String
        Try
            myResponse = Response
            myInsps = OutlookApp.Inspectors
            If myInsps.Count = 0 Then
                ' the user hit Reply from the Explorer window -- there's not an item open in an Inspector window
                ' 11/11/2015 changed myOriginal = OutlookApp.ActiveExplorer.Selection.Item(1)
                myExpl = OutlookApp.ActiveExplorer
                mySel = myExpl.Selection
                myOriginal = mySel.Item(1)
                Marshal.ReleaseComObject(mySel)
                Marshal.ReleaseComObject(myExpl)
            Else
                ' For Each myInsp In OutlookApp.Inspectors
                myInsp = myInsps(1)
                myOriginal = myInsp.CurrentItem
                If TypeOf myOriginal Is Outlook.MailItem Then
                Else
                    MsgBox("The item in the first Outlook Inspector window is not a MailItem.", vbOKOnly + vbExclamation, "ReplyOrReplyAll()")
                    Return
                End If
                Marshal.ReleaseComObject(myInsp)
            End If
            Marshal.ReleaseComObject(myInsps)

            ' 11/4/2015 Replying to an email that was Forwarded to you won't work without this
            str1 = RemoveREFW(myOriginal.Subject)
            str2 = RemoveREFW(myResponse.Subject)

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
                    Marshal.ReleaseComObject(myAttachment)
                Next
                ' this is not in the Access code -- it's used to keep track of whether or not the email originated in InstantFile or Outlook
                myProperties = myResponse.UserProperties
                myProp = myProperties.Add("CameFromOutlook", Outlook.OlUserPropertyType.olText)
                myProp.Value = strEventName
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "ReplyOrReplyAll()")
        Finally
            If myProp IsNot Nothing Then Marshal.ReleaseComObject(myProp) : myProp = Nothing
            If myProperties IsNot Nothing Then Marshal.ReleaseComObject(myProperties) : myProperties = Nothing
            If myInsp IsNot Nothing Then Marshal.ReleaseComObject(myInsp) : myInsp = Nothing
            If mySel IsNot Nothing Then Marshal.ReleaseComObject(mySel) : mySel = Nothing
            If myExpl IsNot Nothing Then Marshal.ReleaseComObject(myExpl) : myExpl = Nothing
            If myInsps IsNot Nothing Then Marshal.ReleaseComObject(myInsps) : myInsps = Nothing
            If myOriginal IsNot Nothing Then Marshal.ReleaseComObject(myOriginal) : myOriginal = Nothing
            If myResponse IsNot Nothing Then Marshal.ReleaseComObject(myResponse) : myResponse = Nothing
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
        ' Debug.Print("ProcessBeforeAttachmentPreview")
    End Sub

    Public Overrides Sub ProcessBeforeAttachmentRead(ByVal attachment As Object, ByVal e As AddinExpress.MSO.ADXCancelEventArgs)
        Const strMsg As String = "This will only work if InstantFile is open." & vbNewLine & vbNewLine & _
                                 "Open InstantFile, then try this again."
        Dim myAttachment As Outlook.Attachment = Nothing
        Dim appAccess As Access.Application = Nothing
        Dim myNote As Outlook.NoteItem = Nothing
        Dim olNameSpace As Outlook.NameSpace = Nothing
        Dim olItem As Object = Nothing
        Try
            myAttachment = attachment
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
                    End Try
                    e.Cancel = True
                    Return
                End If
            ElseIf Left(myAttachment.DisplayName, Len(strIFtaskTag)) = strIFtaskTag Then  ' added 11/16/2015
                Dim strFileName As String
                With myAttachment
                    strFileName = "C:\tmp\" & .FileName
                    .SaveAsFile(strFileName)
                End With
                myNote = OutlookApp.CreateItemFromTemplate(strFileName)
                Dim strID As String, x As Short
                strID = Mid(myNote.Body, Len(strIFtaskTag) + 3)
                x = InStr(1, strID, vbNewLine)
                strID = Left(strID, x - 1)
                myNote.Close(Outlook.OlInspectorClose.olDiscard)
                Marshal.ReleaseComObject(myNote)
                Try
                    olNameSpace = OutlookApp.GetNamespace("MAPI")
                    olItem = olNameSpace.GetItemFromID(strID)  ' couldn't get this to work with the StoreID, but it works without the 2nd argument
                    olItem.Display()
                Catch ex As Exception
                    MsgBox("The InstantFile Request could not be displayed.", vbExclamation, "Display InstantFile Note")
                End Try
                e.Cancel = True
            End If
        Finally
            If olItem IsNot Nothing Then Marshal.ReleaseComObject(olItem) : olItem = Nothing
            If olNameSpace IsNot Nothing Then Marshal.ReleaseComObject(olNameSpace) : olNameSpace = Nothing
            If myNote IsNot Nothing Then Marshal.ReleaseComObject(myNote) : myNote = Nothing
            If appAccess IsNot Nothing Then Marshal.ReleaseComObject(appAccess) : appAccess = Nothing
            ' myAttachment refers to object that was passed into procedure, so don't release it
        End Try
    End Sub

    Public Overrides Sub ProcessBeforeAttachmentWriteToTempFile(ByVal Attachment As Object, ByVal E As AddinExpress.MSO.ADXCancelEventArgs)
        'Debug.Print("ProcessBeforeAttachmentPreview fired.")
    End Sub

    Public Overrides Sub ProcessUnload()
        ' Debug.Print("ProcessUnload() fired")
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
        ' updated this 11/16/2015
        Dim strDisplayName As String
        Dim intX As Integer
        Try
            If Left(myAttach.DisplayName, 18) = strIFmatNo Then
                strDisplayName = Mid(myAttach.DisplayName, 19)
                intX = InStr(1, strDisplayName, Space(1))
                If intX > 0 Then strDisplayName = Left(strDisplayName, intX - 1)
                Return CDbl(strDisplayName)
            ElseIf Left(myAttach.DisplayName, 18) = strIFdocNo Then
                Return CDbl(MatNoFromSubject(strSubject))
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(Err.Description, vbExclamation, "Parse MatterNo from Attachment")
            Return False
        End Try
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


