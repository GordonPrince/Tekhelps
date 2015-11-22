Imports System.Data  ' includes SqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Diagnostics

Module Globals
    Public Const strDocScanned As String = "Document scanned + imported:"
    Public Const strLastScanned As String = "LAST REQUESTED DOCUMENT scanned + imported:"
    Public Const strIFmatNo As String = "InstantFile_MatNo_"
    Public Const strIFdocNo As String = "InstantFile_DocNo_"
    Public Const strPublicFolders As String = "Public Folders"
    Public Const strAllPublicFolders As String = "All Public Folders"
    Public Const strInstantFile As String = "InstantFile"
    Public Const strNewCallTrackingTag As String = "NewCall Tracking Item"
    Public Const strNewCallAppointmentTag As String = "NewCall Appointment"
    Public Const strIFtaskTag As String = "InstantFile_Task"
    Public Const strSend2Gordon As String = vbNewLine & vbNewLine & "Please use the Snipping Tool to capture this message" & vbNewLine & _
                                                "and E-mail it to Gordon."

    Public strScratch As String
    Public strPublicStoreID As String
    Public RetVal As VariantType
    Public lngX As Long

    Public myPublicFolder As Outlook.Folder = Nothing

    ' Private OutlookApp As Outlook.Application = CType(AddinModule.CurrentInstance, AddinModule).OutlookApp
    Public OutlookApp As Outlook.Application = Nothing

    Public Function RunSQLcommand(ByVal queryString As String) As Boolean
        Dim strConnectionString As String = SQLConnectionString()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        Dim cmd As New SqlClient.SqlCommand(queryString, con)
        ' Using con As New SqlClient.SqlConnection(strConnectionString)
        Try
            cmd.Connection.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
        ' End Using
    End Function

    Public Function SQLConnectionString() As String
        If My.Computer.Name = "TEKHELPS7X64" Then
            SQLConnectionString = ("Initial Catalog=InstantFile;Data Source=TEKHELPS7X64\SQL2005X64;Integrated Security=SSPI;")
        Else
            SQLConnectionString = ("Initial Catalog=InstantFile;Data Source=SQLserver;Integrated Security=SSPI;")
        End If
    End Function

    Public Function GetPublicFolder(ByVal strFolderName As String) As Boolean ' 11/17/2015
        Dim mySession As Outlook.NameSpace = Nothing
        Dim myFolders As Outlook.Folders = Nothing
        Dim myFolder As Outlook.Folder = Nothing

        ' also try myFolder = myFolders.GetNext 
        ' https://msdn.microsoft.com/en-us/library/office/ff865587.aspx?f=255&MSPPError=-2147217396
        Try
            mySession = OutlookApp.Session
            myFolders = mySession.Folders
            If myFolders.Count > 0 Then
                Dim x As Short
                For x = 1 To myFolders.Count
                    myFolder = myFolders(x)
                    If Left(myFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
                        Marshal.ReleaseComObject(myFolders)
                        myFolders = myFolder.Folders
                        Marshal.ReleaseComObject(myFolder)
                        Dim y As Short
                        For y = 1 To myFolders.Count
                            myFolder = myFolders(y)
                            If myFolder.Name = strAllPublicFolders Then
                                'Marshal.ReleaseComObject(myFolders)
                                myFolders = myFolder.Folders
                                Marshal.ReleaseComObject(myFolder)
                                myFolder = myFolders.GetFirst
                                Do While Not myFolder Is Nothing
                                    If myFolder.Name = strFolderName Then
                                        If myPublicFolder IsNot Nothing Then Marshal.ReleaseComObject(myPublicFolder)
                                        myPublicFolder = myFolder
                                        Return True
                                    Else
                                        myFolder = myFolders.GetNext
                                    End If
                                Loop
                                Return False
                            End If
                        Next ' Public Folders
                    End If
                Next ' Session folders
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "GetPublicFolder()")
        Finally
            ' don't release myFolder -- it's the same object as myPublicFolder
            ' If myFolder IsNot Nothing Then Marshal.ReleaseComObject(myFolder) : myFolder = Nothing
            If myFolders IsNot Nothing Then Marshal.ReleaseComObject(myFolders) : myFolders = Nothing
            If mySession IsNot Nothing Then Marshal.ReleaseComObject(mySession) : mySession = Nothing
        End Try
        Return False
    End Function

    Function EmailMatNo(myAttach As Outlook.Attachment, strSubject As String) As Double

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
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "Parse MatterNo from Attachment")
            Return 0
        End Try
    End Function

    Function MatNoFromSubject(ByVal strSubject) As Double
        ' try to parse the MatterNo from the Subject line, not the attachment
        Dim intA As Integer, intB As Integer
        Dim strSearchFor As String = vbNullString

        ' check for either string in the Subject. Use whichever one is found
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
                Try
                    Return CDbl(Left(strSubject, intB))
                Catch ex As Exception
                    Return 0
                End Try
            Else
                Return 0
            End If
        Else
            Return 0
        End If
    End Function

    Function InterceptNote(attachment) As Boolean
        Const strMsg As String = "This will only work if InstantFile is open." & vbNewLine & vbNewLine & _
                                 "Open InstantFile, then try this again."
        Const strTitle As String = "InterceptNote() for "
        Dim myAttachment As Outlook.Attachment = Nothing
        Dim appAccess As Access.Application = Nothing
        Dim myNote As Outlook.NoteItem = Nothing
        Dim olNameSpace As Outlook.NameSpace = Nothing
        Dim olItem As Object = Nothing
        Debug.Print("InterceptNote fired")
        Try
            myAttachment = attachment
            If Left(myAttachment.DisplayName, Len(strIFdocNo)) = strIFdocNo Then
                Const strDoc As String = "Open InstantFile Document"
                Dim lngDocNo As Long = Mid(myAttachment.DisplayName, 19)
                If IsDBNull(lngDocNo) Or lngDocNo = 0 Then
                    MsgBox("The item does not have a DocNo.", vbExclamation, strTitle & strDoc)
                    Return False
                Else
                    Try
                        appAccess = CType(Marshal.GetActiveObject("Access.Application"), Access.Application)
                        If Not appAccess.Visible Then appAccess.Visible = True
                        appAccess.Run("DisplayDocument", lngDocNo)
                        Debug.Print("InterceptNote: DisplayDocument")
                        Marshal.ReleaseComObject(appAccess)
                        Return True
                    Catch
                        MsgBox(strMsg, vbExclamation + vbOKOnly, strTitle & strDoc)
                    End Try
                    Return False
                End If
            ElseIf Left(myAttachment.DisplayName, Len(strIFmatNo)) = strIFmatNo Then
                Const strMat As String = "Show Matter in InstantFile"
                Dim dblMatNo As Double = Mid(myAttachment.DisplayName, 19)
                If IsDBNull(dblMatNo) Or dblMatNo = 0 Then
                    MsgBox("The item does not have a MatterNo.", vbExclamation, strTitle & strMat)
                    Return False
                Else
                    Try
                        appAccess = CType(Marshal.GetActiveObject("Access.Application"), Access.Application)
                        If Not appAccess.Visible Then appAccess.Visible = True
                        appAccess.Run("DisplayMatter", dblMatNo)
                        Marshal.ReleaseComObject(appAccess)
                        Return True
                    Catch
                        MsgBox(strMsg, vbExclamation + vbOKOnly, strTitle & strMat)
                    End Try
                    Debug.Print("InterceptNote: DisplayMatter")
                    Return False
                End If
            ElseIf Left(myAttachment.DisplayName, Len(strIFtaskTag)) = strIFtaskTag Then  ' added 11/16/2015, updated 11/17/2015
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
                'strID = Mid(myNote.Body, Len(strIFtaskTag) + 1) ' strip out the tag
                'x = InStr(1, strID, vbNewLine)
                'If x > 0 Then strID = Mid(strID, x + 2) ' strip out the leading vbNewLine, which should leave only the EntryID

                myNote.Close(Outlook.OlInspectorClose.olDiscard)
                Marshal.ReleaseComObject(myNote)
                Try
                    olNameSpace = OutlookApp.GetNamespace("MAPI")
                    olItem = olNameSpace.GetItemFromID(strID)
                    olItem.Display()
                    Return True
                Catch ex As Exception
                    MsgBox("The InstantFile Request could not be displayed.", vbExclamation, strTitle & strIFtaskTag)
                End Try
                Return False
            ElseIf Left(myAttachment.DisplayName, Len(strNewCallTrackingTag)) = strNewCallTrackingTag Then  ' added 11/17/2015
                Dim strFileName As String
                With myAttachment
                    strFileName = "C:\tmp\" & .FileName
                    .SaveAsFile(strFileName)
                End With
                myNote = OutlookApp.CreateItemFromTemplate(strFileName)
                Dim strID As String, x As Short
                strID = Mid(myNote.Body, Len(strNewCallTrackingTag) + 1) ' strip out the tag
                x = InStr(1, strID, vbNewLine)
                If x > 0 Then strID = Mid(strID, x + 2) ' strip out the leading vbNewLine, which should leave only the EntryID
                myNote.Close(Outlook.OlInspectorClose.olDiscard)
                Marshal.ReleaseComObject(myNote)
                Try
                    olNameSpace = OutlookApp.GetNamespace("MAPI")
                    olItem = olNameSpace.GetItemFromID(strID)
                    olItem.Display()
                    Return True
                Catch ex As Exception
                    MsgBox("The " & strNewCallTrackingTag & " could not be displayed.", vbExclamation, strTitle & strNewCallTrackingTag)
                End Try
                Return False
            ElseIf Left(myAttachment.DisplayName, Len(strNewCallAppointmentTag)) = strNewCallAppointmentTag Then  ' added 11/17/2015
                Dim strFileName As String
                With myAttachment
                    strFileName = "C:\tmp\" & .FileName
                    .SaveAsFile(strFileName)
                End With
                myNote = OutlookApp.CreateItemFromTemplate(strFileName)
                Dim strID As String, x As Short
                strID = Mid(myNote.Body, Len(strNewCallAppointmentTag) + 1) ' strip out the tag
                x = InStr(1, strID, vbNewLine)
                If x > 0 Then strID = Mid(strID, x + 2) ' strip out the leading vbNewLine, which should leave only the EntryID
                myNote.Close(Outlook.OlInspectorClose.olDiscard)
                Marshal.ReleaseComObject(myNote)
                Try
                    olNameSpace = OutlookApp.GetNamespace("MAPI")
                    olItem = olNameSpace.GetItemFromID(strID)
                    olItem.Display()
                    Return True
                Catch ex As Exception
                    MsgBox("The " & strNewCallAppointmentTag & " could not be displayed.", vbExclamation, strTitle & strNewCallAppointmentTag)
                End Try
                Return False
            Else
                Return False
            End If
        Finally
            If olItem IsNot Nothing Then Marshal.ReleaseComObject(olItem) : olItem = Nothing
            If olNameSpace IsNot Nothing Then Marshal.ReleaseComObject(olNameSpace) : olNameSpace = Nothing
            If myNote IsNot Nothing Then Marshal.ReleaseComObject(myNote) : myNote = Nothing
            If appAccess IsNot Nothing Then Marshal.ReleaseComObject(appAccess) : appAccess = Nothing
            ' myAttachment refers to object that was passed into procedure, so don't release it
        End Try
    End Function
End Module
