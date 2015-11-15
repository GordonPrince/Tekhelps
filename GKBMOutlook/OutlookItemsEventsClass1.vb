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
 
    Dim OutlookApp As Outlook.Application = CType(AddinModule.CurrentInstance, AddinModule).OutlookApp

    Public Sub New(ByVal ADXModule As AddinExpress.MSO.ADXAddinModule)
        MyBase.New(ADXModule)
    End Sub

    Public Overrides Sub ItemAdd(ByVal Item As Object, ByVal SourceFolder As Object)
        Const strTitle As String = "Save Sent E-mail in InstantFile"
        Const strCommentID As String = "InstantFile CommentID "
        Const strDocNo As String = "InstantFile DocNo "

        Dim myFolder As Outlook.MAPIFolder = Nothing
        Dim myMailItem As Outlook.MailItem = Nothing
        Dim myCopy As Outlook.MailItem = Nothing
        Dim myMove As Outlook.MailItem = Nothing
        Dim myAttachments As Outlook.Attachments = Nothing
        Dim myAttach As Outlook.Attachment = Nothing
        Dim myProperties As Outlook.UserProperties = Nothing
        Dim myProp As Outlook.UserProperty = Nothing
        Dim myRecipients As Outlook.Recipients = Nothing
        Dim myRecipient As Outlook.Recipient = Nothing

        Dim strSQL As String, strBody As String = "email to "
        Dim dblMatNo As Double, intA As Integer, lngDocNo As Long
        Dim bScanned As Boolean

        Static strLastID As String

        'Dim myFolder As Outlook.Folder = SourceFolder
        'Debug.Print("ItemAdd() myFolder.FolderPath = " & myFolder.FolderPath)

        If TypeOf Item Is Outlook.MailItem Then
            myMailItem = Item
            'Debug.Print("myMailItem.Subject = " & myMailItem.Subject)
        Else
            Return
        End If

        ' Outlook seems to process each item twice. The first time works, subsequent times fail
        Try
            If myMailItem.EntryID = strLastID Then
                Exit Sub
            Else
                strLastID = myMailItem.EntryID
            End If
        Catch ex As Exception
            Return
        End Try

        If Left(myMailItem.Subject, 13) = "Task Request:" _
        Or Left(myMailItem.Subject, 14) = "Task Accepted:" _
        Or Left(myMailItem.Subject, 14) = "Task Declined:" Then
            ' it was created by InstantFile, therefore it's already been stored in InstantFile
            Return
        End If

        Try
            ' save Sent MailItems as comments if they have the attachment that Import2InstantFile creates
            If Len(myMailItem.BillingInformation) > 0 Then
                If InStr(1, myMailItem.BillingInformation, strCommentID) Or InStr(1, myMailItem.BillingInformation, strDocNo) Then
                    GoTo InstantFileEmail
                End If
            ElseIf Len(myMailItem.Subject) > 0 Then
                If Left(myMailItem.Subject, Len(strDocScanned)) = strDocScanned Or Left(myMailItem.Subject, Len(strLastScanned)) = strLastScanned Then
                    bScanned = True
                    GoTo InstantFileEmail
                End If
            Else
                ' if this is an InstantFile related E-mail then add it to InstantFile (unless it originated in InstantFile)
                'For Each myAttach In myMailItem.Attachments
                Dim x As Short
                myAttachments = myMailItem.Attachments
                For x = 1 To myAttachments.Count
                    myAttach = myAttachments(x)
                    dblMatNo = EmailMatNo(myAttach, myMailItem.Subject)
                    If dblMatNo > 0 Then
                        myProperties = myMailItem.UserProperties
                        ' myProp = myMailItem.UserProperties.Find("CameFromOutlook")
                        myProp = myProperties.Find("CameFromOutlook")
                        If myProp Is Nothing Then Return
                        If MsgBox("Save the E-mail you sent as an InstantFile Comment in matter " & dblMatNo & "?", vbQuestion + vbYesNo, strTitle) = vbYes Then
                            bScanned = False
                            If dblMatNo > 0 Then
                                GoTo InstantFileEmail
                            Else
                                dblMatNo = InputBox("Enter the Matter # to save this comment under", strTitle, "0.00")
                                If dblMatNo = 0 Then
                                    MsgBox("No comment was added to InstantFile about this E-mail.", vbInformation, strTitle)
                                    Return
                                End If
                            End If
                        Else
                            Return
                        End If
                        Marshal.ReleaseComObject(myProp)
                        Marshal.ReleaseComObject(myProperties)
                    End If
                    Marshal.ReleaseComObject(myAttach)
                Next

                ' if no Note attachment had the MatterNo on it, try to determine the MatterNo from the DocNo that's attached
                ' For Each myAttach In myMailItem.Attachments
                For x = 1 To myAttachments.Count
                    myAttach = myAttachments(x)
                    strScratch = myAttach.DisplayName
                    ' added 10/25/2010
                    If strScratch = strNewCallTrackingTag Then
                    Else
                        intA = InStr(1, strScratch, strIFdocNo)
                        If intA Then
                            Try
                                lngDocNo = Mid(strScratch, intA + Len(strIFdocNo))
                            Catch
                                MsgBox("Could not set lngDocNo from Mid(strScratch, intA + Len(strIFdocNo))" & vbNewLine & _
                                        strScratch & strSend2Gordon, vbInformation, strTitle)
                                Exit Sub
                            End Try
                        End If
                        If lngDocNo > 0 Then
                            '    With con
                            '        .Open(strConnectionString)
                            '        rst = .Execute("sp_MatNo4DocNo " & lngDocNo)
                            '    End With
                            '    With rst
                            '        If .EOF Then
                            '            MsgBox("Could not find MatterNo for DocNo=" & lngDocNo & "." & vbNewLine & vbNewLine & _
                            '                    "Please forward the E-mail you just sent to Gordon" & vbNewLine & _
                            '                    "and type 'Could not find MatterNo for DocNo' as the subject.", vbExclamation, strTitle)
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
                    Marshal.ReleaseComObject(myAttach)
                Next
                Marshal.ReleaseComObject(myAttachments)
            End If

            ' if you get here there either aren't any attachments or 
            ' it's not an Import2InstantFile document that's attached or 
            ' the attachment is a NewCallTracking note
            Return

InstantFileEmail:
            ' don't know why this wouldn't work
            ' For Each pFolder In OutlookApp.Session.Folders   
            Const strFolderName As String = "InstantFile Mail"
            If Not GetPublicFolder(strFolderName, myFolder) Then
                MsgBox("Could not find the Public Folder '" & strFolderName & "'", vbExclamation, strTitle)
                Return
            End If

            myCopy = myMailItem.Copy
            myMove = myCopy.Move(myFolder)  ' the myMove object has the new EntryID

            Dim con As SqlClient.SqlConnection, myCmd As SqlClient.SqlCommand, myReader As SqlClient.SqlDataReader
            Dim strInitials As String = Nothing
            con = New SqlClient.SqlConnection(SQLConnectionString)
            myCmd = con.CreateCommand
            strScratch = "select staff_id from staff where staff_name = '" & Replace(mySession.CurrentUser.Name, "'", "''") & "'"
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
                        "Please tell Gordon this message appeared and have him make your Outlook user name and your InstantFile user name the same.", vbExclamation, strTitle)
                strInitials = InputBox("Enter your initials", strTitle, "ABC")
                If Len(strInitials) > 0 Then
                Else
                    MsgBox("No initials were entered. No comment about this E-mail could be created.", vbExclamation, strTitle)
                    con.Close()
                    Exit Sub
                End If
            End If

            Dim lngX As Int16
            With myMove
                ' This updates the database with the EntryID of the mail item, which can only be done after it was sent
                ' InstantFile puts the Comment or Email.DocNo in the BillingInformation field when it creates the email
                ' So this code just needs to parse the ID from there and then update the database with the EntryID
                ' Debug.Print("ItemAdd() myMove.BillingInformation = " & .BillingInformation)
                If InStr(1, .BillingInformation, strCommentID) Or InStr(1, .BillingInformation, strDocNo) Then
                    ' update the Comment with the EntryID
                    lngX = InStr(1, .BillingInformation, strCommentID)
                    If lngX > 0 Then
                        strSQL = Mid(.BillingInformation, Len(strCommentID))
                        lngX = InStr(1, strSQL, ",")
                        If lngX > 0 Then strSQL = Left(strSQL, lngX - 1)
                        strSQL = Trim(strSQL)
                        strScratch = "UPDATE COMMENT SET EntryID = '" & .EntryID & "' WHERE CommentID = " & CLng(strSQL)
                        If Not RunSQLcommand(strScratch) Then
                            MsgBox("The InstantFile Comment was not updated properly with the E-mail's EntryID.", vbExclamation, strTitle)
                        End If
                    End If
                    ' update the Email row with the EntryID
                    lngX = InStr(1, .BillingInformation, strDocNo)
                    If lngX > 0 Then
                        strSQL = Mid(.BillingInformation, lngX + 1)
                        strSQL = Trim(Mid(strSQL, Len(strDocNo)))
                        strScratch = "UPDATE Email SET EntryID = '" & .EntryID & "' WHERE DocNo = " & CLng(strSQL)
                        If Not RunSQLcommand(strScratch) Then
                            MsgBox("The InstantFile Document was not updated properly with the E-mail's EntryID.", vbExclamation, strTitle)
                        End If
                    End If
                    ' Debug.WriteLine("The E-mail's EntryID was updated in InstantFile.")
                    ' without the MsgBox here I get an error
                    ' MsgBox("The E-mail's EntryID was updated in InstantFile.", vbInformation + vbOKOnly, "GKBM Outlook Add-in")
                    Exit Sub
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
                    dblMatNo = InputBox("Enter the Matter # this E-mail should be saved in.", strTitle)
                End If

AddRecipientsAndBody:
                strBody = "email to "
                ' For Each myRecipient In .Recipients
                Dim r As Short
                myRecipients = .Recipients
                For r = 1 To myRecipients.Count
                    myRecipient = myRecipients(r)
                    strBody = strBody & myRecipient.Name & "; "
                    Marshal.ReleaseComObject(myRecipient)
                Next ' myRecipient
                Marshal.ReleaseComObject(myRecipients)
                strBody = Left(strBody, Len(strBody) - 2)

                If bScanned Then
                    strBody = strBody & strDocScanned
                    intA = InStr(Len(strDocScanned) + 1, .Body, "Author:")
                    If intA > 0 Then strBody = strBody & Trim(Mid(.Body, intA + 7))
                Else
                    strBody = strBody & " -- " & myMove.Body
                End If
            End With

            If Len(strBody) > 0 Then
                strBody = Replace(strBody, "Summary:", vbNullString)
                strBody = LTrim(strBody)
                Do While InStr(1, strBody, vbNewLine & vbNewLine)
                    strBody = Replace(strBody, vbNewLine & vbNewLine, vbNewLine)
                Loop
            End If

            strSQL = "INSERT INTO COMMENT (matter_no, author, summary, EntryID)" & _
                    " VALUES (" & dblMatNo & ",'" & strInitials & "','" & Left(Replace(strBody, "'", "''"), 2000) & "','" & myMove.EntryID & "')"
            If RunSQLcommand(strSQL) Then
                MsgBox("An InstantFile Comment was created from your E-mail" & vbNewLine & _
                       "and a copy of the E-mail was saved with the Comment.", vbInformation, strTitle)
            End If

        Catch ex As Exception
            MsgBox(Err.Description, vbExclamation, strTitle)
        Finally
            If myRecipient IsNot Nothing Then Marshal.ReleaseComObject(myRecipient) : myRecipient = Nothing
            If myRecipients IsNot Nothing Then Marshal.ReleaseComObject(myRecipients) : myRecipients = Nothing
            If myProp IsNot Nothing Then Marshal.ReleaseComObject(myProp) : myProp = Nothing
            If myProperties IsNot Nothing Then Marshal.ReleaseComObject(myProperties) : myProperties = Nothing
            If myAttach IsNot Nothing Then Marshal.ReleaseComObject(myAttach) : myAttach = Nothing
            If myAttachments IsNot Nothing Then Marshal.ReleaseComObject(myAttachments) : myAttachments = Nothing
            If myMove IsNot Nothing Then Marshal.ReleaseComObject(myMove) : myMove = Nothing
            If myCopy IsNot Nothing Then Marshal.ReleaseComObject(myCopy) : myCopy = Nothing
            If myMailItem IsNot Nothing Then Marshal.ReleaseComObject(myMailItem) : myMailItem = Nothing
            If myFolder IsNot Nothing Then Marshal.ReleaseComObject(myFolder) : myFolder = Nothing
        End Try

    End Sub

    Function EmailMatNo(myAttach As Outlook.Attachment, strSubject As String) As Double
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
            EmailMatNo = 0
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

    Public Overrides Sub ItemChange(ByVal Item As Object, ByVal SourceFolder As Object)
        ' Debug.Print("ItemChange() fired")
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

End Class

