Imports System
Imports System.Windows.Forms
Imports AddinExpress.MSO
Imports System.Diagnostics
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
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
    Const strPublicFolders As String = "Public Folders"

    Public Overrides Sub ItemAdd(ByVal Item As Object, ByVal SourceFolder As Object)
        ' MsgBox("ItemAdd fired.")
        Const strTitle As String = "Save Sent E-mail in InstantFile"
        Const strFolderName As String = "All Public Folders"
        Const strCommentID As String = "InstantFile CommentID "
        Const strDocNo As String = "InstantFile DocNo "


        Dim pFolder As Outlook.MAPIFolder, aFolder As Outlook.MAPIFolder, mFolder As Outlook.MAPIFolder
        Dim myMailItem As Outlook.MailItem, myCopy As Outlook.MailItem, myMove As Outlook.MailItem
        Dim myAttachment As Outlook.Attachment
        Dim myRecipient As Outlook.Recipient
        Dim strSQL As String, strBody As String

        Dim dblMatNo As Double, intA As Integer, intB As Integer, lngDocNo As Long
        Dim bScanned As Boolean, myUserProp As Outlook.UserProperty

        Static strLastID As String

        Dim appOutlook As Outlook.Application = Nothing
        If TypeOf Item Is Outlook.MailItem Then
            myMailItem = Item
            appOutlook = myMailItem.Application
        End If

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
            For Each myAttachment In myMailItem.Attachments
                If TypeOf myAttachment.Application Is Outlook.Application And myAttachment.Class = 5 Then
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
                ' If myAttachment.Application = "Outlook" And myAttachment.Class = 5 Then
                If TypeOf myAttachment.Application Is Outlook.Application And myAttachment.Class = 5 Then
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
                            '    With con
                            '        .Open(strConnectionString)
                            '        rst = .Execute("sp_MatNo4DocNo " & lngDocNo)
                            '    End With
                            '    With rst
                            '        If .EOF Then
                            '            MsgBox("Could not find MatterNo for DocNo=" & lngDocNo & "." & vbNewLine & vbNewLine & _
                            '                    "Please forward the email you just sent to Gordon" & vbNewLine & _
                            '                    "and type 'Could not find MatterNo for DocNo' as the message body.", vbExclamation, strTitle)
                            '        Else
                            '            dblMatNo = .Fields("matter_no")
                            '        End If
                            '        .Close()
                            '    End With
                            '    rst = Nothing
                        End If
                        ' if dblmatno is not set, prompt for the MatterNo after prompting to save the email
                        ' GoTo Prompt2Save
                    End If
                End If
            Next
        End If

DontAdd2InstantFile:
        ' if you get here there either aren't any attachments 
        ' or it's not an Import2InstantFile document that's attached 
        ' or the attachment is a NewCallTracking note
        Exit Sub

InstantFileEmail:
        For Each pFolder In appOutlook.Session.Folders
            If Left(pFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
                For Each aFolder In pFolder.Folders
                    If aFolder.Name = "All Public Folders" Then
                        For Each mFolder In aFolder.Folders
                            If mFolder.Name = "InstantFile Mail" Then GoTo HaveInstantFileMailFolder
                        Next
                    End If
                Next
            End If
        Next
        MsgBox("Could not find the folder 'InstantFile Mail'", vbExclamation, strTitle)
        Exit Sub

HaveInstantFileMailFolder:
        myCopy = myMailItem.Copy
        myMove = myCopy.Move(mFolder)  ' the myMove object has the new EntryID
        ' MsgBox("A copy of the item was saved in the InstantFile Mail folder.", vbInformation, strTitle)

        ' make a comment in InstantFile about this email
        'With con
        '    If .State = 0 Then .Open(strConnectionString)
        '    rst = .Execute("select staff_id from staff where staff_name = '" & Replace(Outlook.Session.CurrentUser, "'", "''") & "'") ' in case there's an apostrophe in the staff_name
        'End With
        ' Dim con As New ADODB.Connection, rst As ADODB.Recordset

        '"Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=C:\Access\Access2010\GKBM\OutlookStubs.accdb;" & _
        '"Mode=Share Deny None;Extended Properties="";" & _
        '"Jet OLEDB:System database=C:\Users\Gordon\AppData\Roaming\Microsoft\Access\System.mdw;" & _
        '"Jet OLEDB:Registry Path=Software\Microsoft\Office\14.0\Access\Access Connectivity Engine;" & _
        '"Jet OLEDB:Database Password="";" & _
        '"Jet OLEDB:Engine Type=6;" & _
        '"Jet OLEDB:Database Locking Mode=1;" & _
        '"Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;" & _
        '"Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;" & _
        '"Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;" & _
        '"Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=True;Jet OLEDB:Bypass UserInfo Validation=False"
        ' Const strConnectionString As String = _
        ' "App=GKBMOutlookAdd-in;Provider=MSDataShape.1;Persist Security Info=False;Data Source=SQLserver;Integrated Security=SSPI;" & _
        ' "Initial Catalog=InstantFile;Data Provider=SQLOLEDB.1"
        Dim strConnectionString As String
        If My.Computer.Name = "TEKHELPS7X64" Then
            strConnectionString = ("Initial Catalog=InstantFile;Data Source=TEKHELPS7X64\SQL2005X64;Integrated Security=SSPI;")
        Else
            strConnectionString = ("Initial Catalog=InstantFile;Data Source=TEKHELPS7X64\SQL2005X64;Integrated Security=SSPI;")
        End If

        '"App=GKBMOutlookAdd-in;Data Source=SQLserver;" & _
        '"Database Password=""lahave$13"";" & _
        '"Initial Catalog=InstantFile;"
        Dim con As SqlClient.SqlConnection, myCmd As SqlClient.SqlCommand, myReader As SqlClient.SqlDataReader
        Dim strInitials As String = Nothing
        con = New SqlClient.SqlConnection(strConnectionString)
        myCmd = con.CreateCommand
        strScratch = "select staff_id from staff where staff_name = '" & Replace(appOutlook.Session.CurrentUser.Name, "'", "''") & "'"
        myCmd.CommandText = strScratch
        con.Open()
        myReader = myCmd.ExecuteReader()
        Do While myReader.Read
            strInitials = myReader.GetString(0)
        Loop
        myReader.Close()
        con.Close()

        If Len(strInitials) > 0 Then
        Else
            MsgBox("Could not find your initials based on your Outlook user name." & vbNewLine & vbNewLine & _
                    "Please tell Gordon this message appeared and have him make them the same.", vbExclamation, strTitle)
            strInitials = InputBox("Enter your initials", strTitle, "ABC")
            If Len(strInitials) > 0 Then
            Else
                MsgBox("No initials were entered. No comment about this E-mail could be created.", vbExclamation, strTitle)
                con.Close()
                GoTo SentItems_Exit
            End If
        End If

        Dim lngX As Int16
        With myMove
            If InStr(1, .BillingInformation, strCommentID) Or InStr(1, .BillingInformation, strDocNo) Then
                ' update the Comment with the EntryID
                lngX = InStr(1, .BillingInformation, strCommentID)
                If lngX > 0 Then
                    strSQL = Mid(.BillingInformation, Len(strCommentID))
                    lngX = InStr(1, strSQL, ",")
                    If lngX > 0 Then strSQL = Left(strSQL, lngX - 1)
                    strSQL = Trim(strSQL)
                    strScratch = "update Comment set EntryID = '" & .EntryID & "' where CommentID = " & CLng(strSQL)
                    ' con.Execute(strScratch, lngX)
                    myCmd.CommandText = strScratch
                    con.Open()
                    con.Close()
                    If lngX <> 1 Then MsgBox("The InstantFile Comment was not updated properly with the email's EntryID.", vbExclamation, strTitle)
                End If
                ' update the Email with the EntryID
                lngX = InStr(1, .BillingInformation, strDocNo)
                If lngX > 0 Then
                    strSQL = Mid(.BillingInformation, lngX + 1)
                    strSQL = Trim(Mid(strSQL, Len(strDocNo)))
                    strScratch = "update Email set EntryID = '" & .EntryID & "' where DocNo = " & CLng(strSQL)
                    ' con.Execute(strScratch, lngX)
                    myCmd.CommandText = strScratch
                    con.Open()
                    con.Close()
                    If lngX <> 1 Then MsgBox("The InstantFile Document was not updated properly with the email's EntryID.", vbExclamation, strTitle)
                End If
                GoTo SentItems_Exit
            ElseIf Left(myMailItem.Subject, Len(strDocScanned)) = strDocScanned Then
                intA = InStr(1, Mid(.Subject, Len(strDocScanned) + 2), Space(1))
                If intA > 1 Then
                    dblMatNo = Mid(.Subject, Len(strDocScanned) + 2, intA)
                Else
                    GoTo Prompt4Matter
                End If
            ElseIf Left(myMailItem.Subject, Len(strLastScanned)) = strLastScanned Then
                intA = InStr(1, Mid(.Subject, Len(strLastScanned) + 2), Space(1))
                If intA > 1 Then
                    dblMatNo = Mid(.Subject, Len(strLastScanned) + 2, intA)
                Else
                    GoTo Prompt4Matter
                End If
            ElseIf dblMatNo > 0 Then ' don't prompt for the dblMatNo
            Else
Prompt4Matter:
                dblMatNo = InputBox("Enter the Matter # this email should be saved in.", strTitle)
            End If

            strBody = "email to "
            For Each myRecipient In .Recipients
                strBody = strBody & myRecipient.Name & "; "
            Next myRecipient

            If bScanned Then
                strBody = strBody & strDocScanned
                intA = InStr(Len(strDocScanned) + 1, .Body, "Author:")
                If intA > 0 Then strBody = strBody & Trim(Mid(.Body, intA + 7))
            Else
                strBody = strBody & " -- " & myMove.Body
            End If
        End With

        strBody = Replace(strBody, "Summary:", vbNullString)
        Do While InStr(1, strBody, Chr(160))
            strBody = Replace(strBody, Chr(160), vbNullString)   ' these are spaces
        Loop
        Do While InStr(1, strBody, vbNewLine & vbNewLine)
            strBody = Replace(strBody, vbNewLine & vbNewLine, vbNewLine)
        Loop

        strSQL = "insert into comment (matter_no, author, summary, EntryID)" & _
                " values (" & dblMatNo & ",'" & strInitials & "','" & Left(Replace(strBody, "'", "''"), 2000) & "','" & myMove.EntryID & "')"
        myCmd.CommandText = strSQL
        con.Open()
        myCmd.ExecuteNonQuery()
        con.Close()

        MsgBox("A comment about the email you sent was created in InstantFile" & vbNewLine & _
                "(and a copy of the email was saved with the Comment).", vbInformation, strTitle)

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
        Dim strSearchFor As String = vbNullString

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

