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
 
    Public Overrides Sub ItemAdd(ByVal Item As Object, ByVal SourceFolder As Object)
        MsgBox("ItemAdd fired.")
        Const strTitle As String = "Save Sent E-mail in InstantFile"
        Const strFolderName As String = "All Public Folders"
        Const strCommentID As String = "InstantFile CommentID "
        Const strDocNo As String = "InstantFile DocNo "
        Dim strScratch As String
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
 
End Class

