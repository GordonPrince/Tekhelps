Imports System
Imports System.Windows.Forms
Imports AddinExpress.MSO
Imports System.Diagnostics
Imports Microsoft.Office.Interop
Imports System.Data

'Add-in Express Outlook Items Events Class
Public Class OutlookItemsEventsClass1
    Inherits AddinExpress.MSO.ADXOutlookItemsEvents
 
    Public Sub New(ByVal ADXModule As AddinExpress.MSO.ADXAddinModule)
        MyBase.New(ADXModule)
    End Sub

    Dim strScratch As String
    Const strDocScanned As String = "Document scanned + imported:"
    Const strLastScanned As String = "LAST REQUESTED DOCUMENT scanned + imported:"
    Const strIFmatNo As String = "InstantFile_MatNo_"
    Const strIFdocNo As String = "InstantFile_DocNo_"

    Public Overrides Sub ItemAdd(ByVal Item As Object, ByVal SourceFolder As Object)
        MsgBox("ItemAdd fired.")
        Const strTitle As String = "Save Sent E-mail in InstantFile"
        Const strFolderName As String = "All Public Folders"
        Const strCommentID As String = "InstantFile CommentID "
        Const strDocNo As String = "InstantFile DocNo "

        Const strConnectionString As String = "App=GKBMOutlookAdd-in;Provider=MSDataShape.1;Persist Security Info=False;Data Source=SQLserver;Integrated Security=SSPI;" & _
                                            "Initial Catalog=InstantFile;Data Provider=SQLOLEDB.1"

        Dim pFolder As Outlook.MAPIFolder, aFolder As Outlook.MAPIFolder, mFolder As Outlook.MAPIFolder
        Dim myMailItem As Outlook.MailItem, myCopy As Outlook.MailItem, myMove As Outlook.MailItem
        Dim myAttachment As Outlook.Attachment
        Dim myRecipient As Outlook.Recipient
        Dim strSQL As String, strBody As String
        ' Dim con As New ADODB.Connection, rst As ADODB.Recordset
        Dim con As SqlClient.SqlConnection, myCmd As SqlClient.SqlCommand, rst As SqlClient.SqlDataReader
        Dim varInitials As Object
        Dim dblMatNo As Double, intA As Integer, intB As Integer, lngDocNo As Long
        Dim bScanned As Boolean, myUserProp As Outlook.UserProperty
        Static strLastID As String

        If TypeName(Item) = "MailItem" Then
            If Left(Item.Subject, 13) = "Task Request:" _
            Or Left(Item.Subject, 14) = "Task Accepted:" _
            Or Left(Item.Subject, 14) = "Task Declined:" Then   ' it was created by InstantFile, therefore it's already been stored in InstantFile
                Exit Sub
            End If
            myMailItem = Item
        Else
            Exit Sub
        End If

        ' Outlook 2010 seems to process each item twice. The first time works, subsequent times fail
        On Error Resume Next
        strScratch = myMailItem.EntryID
        If Err.Number = 0 Then
            If myMailItem.EntryID = strLastID Then
                Exit Sub
            Else
                strLastID = myMailItem.EntryID
            End If
        Else
            Err.Clear()
            myMailItem = Nothing
            Exit Sub
        End If
        On Error GoTo SentItems_Error

        ' save Sent MailItems as comments if they have the attachment that Import2InstantFile creates
        ' 2011/11/20 additional logic if not using CDO anymore
        If InStr(1, myMailItem.BillingInformation, strCommentID) Or InStr(1, myMailItem.BillingInformation, strDocNo) Then
            GoTo InstantFileEmail
        ElseIf Left(myMailItem.Subject, Len(strDocScanned)) = strDocScanned Or Left(myMailItem.Subject, Len(strLastScanned)) = strLastScanned Then
            bScanned = True
            GoTo InstantFileEmail
        Else
            ' if this is InstantFile related email then add it to InstantFile (unless it originated in InstantFile)
            'If Environ("username") = "Gordon" Then
            'Else
            For Each myAttachment In myMailItem.Attachments
                If myAttachment.Application = "Outlook" And myAttachment.Class = 5 Then
                    dblMatNo = EmailMatNo(myAttachment, myMailItem.Subject)
                    If dblMatNo > 0 Then
Prompt2Save:
                        myUserProp = myMailItem.UserProperties.Find("CameFromOutlook")
                        ' it didn't come from Outlook as a Reply, ReplyAll or Forward -- it must have come from InstantFile
                        If TypeName(myUserProp) = "Nothing" Then GoTo DontAdd2InstantFile
                        If MsgBox("Save the email you sent as a Comment in matter " & dblMatNo & "?", vbQuestion + vbYesNo, strTitle) = vbYes Then
                            bScanned = False
                            If dblMatNo > 0 Then
                                GoTo InstantFileEmail
                            Else
                                dblMatNo = InputBox("Enter the Matter # to save this comment under", strTitle, "0.00")
                                If dblMatNo = 0 Then
                                    MsgBox("No comment was added to InstantFile about this email.", vbInformation, strTitle)
                                    GoTo DontAdd2InstantFile
                                End If
                            End If
                        Else
                            GoTo DontAdd2InstantFile
                        End If
                    End If
                End If
            Next
            'End If
            ' if no Note attachment had the MatterNo on it, try to determine the MatterNo from the DocNo that's attached
            For Each myAttachment In myMailItem.Attachments
                If myAttachment.Application = "Outlook" And myAttachment.Class = 5 Then
                    strScratch = myAttachment.DisplayName
                    ' added 10/25/2010
                    If strScratch = "NewCall Tracking Item" Then
                    Else
                        intA = InStr(1, strScratch, strIFdocNo)
                        If intA Then
                            On Error Resume Next
                            lngDocNo = Mid(strScratch, intA + Len(strIFdocNo))
                            On Error GoTo SentItems_Error
                        End If
                        If lngDocNo > 0 Then
                            With con
                                .Open(strConnectionString)
                                rst = .Execute("sp_MatNo4DocNo " & lngDocNo)
                            End With
                            With rst
                                If .EOF Then
                                    MsgBox("Could not find MatterNo for DocNo=" & lngDocNo & "." & vbNewLine & vbNewLine & _
                                            "Please forward the email you just sent to Gordon" & vbNewLine & _
                                            "and type 'Could not find MatterNo for DocNo' as the message body.", vbExclamation, strTitle)
                                Else
                                    dblMatNo = .Fields("matter_no")
                                End If
                                .Close()
                            End With
                            rst = Nothing
                        End If
                        ' if dblmatno is not set, prompt for the MatterNo after prompting to save the email
                        GoTo Prompt2Save
                    End If
                End If
            Next
        End If

DontAdd2InstantFile:
        ' if you get here there either aren't any attachments or it's not an Import2InstantFile document that's attached or it's the attachment is a NewCallTracking note
        Exit Sub

InstantFileEmail:


SentItems_Exit:
        Exit Sub

SentItems_Error:
        If Err.Number = 13 Then ' type mismatch
        Else
            MsgBox(Err.Description, vbExclamation, strTitle)
        End If
        GoTo SentItems_Exit
    End Sub
 
    Public Overrides Sub ItemChange(ByVal Item As Object, ByVal SourceFolder As Object)
        'TODO: Add some code
    End Sub
 
    Public Overrides Sub ItemRemove(ByVal SourceFolder As Object)
        'TODO: Add some code
    End Sub
 
    Public Overrides Sub BeforeFolderMove(ByVal moveTo As Object, ByVal SourceFolder As Object, ByVal e As AddinExpress.MSO.ADXCancelEventArgs)
        'TODO: Add some code
    End Sub
 
    Public Overrides Sub BeforeItemMove(ByVal item As Object, ByVal moveTo As Object, ByVal SourceFolder As Object, ByVal e As AddinExpress.MSO.ADXCancelEventArgs)
        'TODO: Add some code
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

    Function MatNoFromSubject(ByVal strSubject) As Double
        ' try to parse the MatterNo from the Subject line, not the attachment
        Dim intA As Integer, intB As Integer
        Dim strSearchFor As String

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

