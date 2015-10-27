Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Windows.Forms
Imports AddinExpress.MSO
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Access = Microsoft.Office.Interop.Access

'Add-in Express Add-in Module
<GuidAttribute("7E29F01B-BDC1-47B5-B1B7-634B70EA309B"), ProgIdAttribute("GKBMOutlook.AddinModule")> _
Public Class AddinModule
    Inherits AddinExpress.MSO.ADXAddinModule

#Region "Tekhelps definitions"
    Const strPublicFolders As String = "Public Folders"
    Const strIFmatNo As String = "InstantFile_MatNo_"
    Const strIFdocNo As String = "InstantFile_DocNo_"
    Const strNewCallTrackingTag As String = "NewCall Tracking Item"
    Const strIFtaskTag As String = "InstantFile_Task"
    Const strNewCallAppointmentTag As String = "NewCall Appointment"
    Dim RetVal As VariantType
    Public strPublicStoreID As String
    Public WithEvents myInspectors As Outlook.Inspectors
    Public WithEvents myInsp As Outlook.Inspector
    Public WithEvents myMailItem As Outlook.MailItem
#End Region

#Region " Add-in Express automatic code "

    'Required by Add-in Express - do not modify
    'the methods within this region

    Public Overrides Function GetContainer() As System.ComponentModel.IContainer
        If components Is Nothing Then
            components = New System.ComponentModel.Container
        End If
        GetContainer = components
    End Function

    <ComRegisterFunctionAttribute()> _
    Public Shared Sub AddinRegister(ByVal t As Type)
        AddinExpress.MSO.ADXAddinModule.ADXRegister(t)
    End Sub

    <ComUnregisterFunctionAttribute()> _
    Public Shared Sub AddinUnregister(ByVal t As Type)
        AddinExpress.MSO.ADXAddinModule.ADXUnregister(t)
    End Sub

    Public Overrides Sub UninstallControls()
        MyBase.UninstallControls()
    End Sub

#End Region

    Public Shared Shadows ReadOnly Property CurrentInstance() As AddinModule
        Get
            Return CType(AddinExpress.MSO.ADXAddinModule.CurrentInstance, AddinModule)
        End Get
    End Property

    Public ReadOnly Property OutlookApp() As Outlook._Application
        Get
            Return CType(HostApplication, Outlook._Application)
        End Get
    End Property

    Private Sub AdxRibbonButton4_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton4.OnClick
        MsgBox("Microsoft Outlook Add-in for" & vbNewLine & _
               "Gatti, Keltner, Bienvenu & Montesi, PLC." & vbNewLine & vbNewLine & _
               "Copyright (c) 1997-2015 by Tekhelps, Inc." & vbNewLine & _
               "For further information contact Gordon Prince (901) 761-3393." & vbNewLine & vbNewLine & _
               "This version dated 2015-Oct-27  7:55.", vbInformation, "About this Add-in")
    End Sub

    Private Sub AdxRibbonButtonSaveAttachments_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButtonSaveAttachments.OnClick
        ' copied from http://www.howto-outlook.com/howto/saveembeddedpictures.htm
        Const strTitle As String = "Save Attachments"
        Dim myOlNameSpace As Outlook.NameSpace, myOlSelection As Outlook.Selection
        Dim mySelectedItem As Object, intPos As Integer
        Dim colAttachments As Outlook.Attachments, objAttachment As Outlook.Attachment
        Dim DateStamp As String, MyFile As String
        Dim intCounter As Integer

        'Get all selected items
        myOlNameSpace = OutlookApp.GetNamespace("MAPI")
        myOlSelection = OutlookApp.ActiveExplorer.Selection
        'Make sure at least one item is selected
        If myOlSelection.Count = 0 Then
            RetVal = MsgBox("Please select an item first.", vbExclamation, strTitle)
            Exit Sub
        End If

        'Make sure only one item is selected
        If myOlSelection.Count > 1 Then
            RetVal = MsgBox("Please select only one item.", vbExclamation, strTitle)
            Exit Sub
        End If

        'Retrieve the selected item
        mySelectedItem = myOlSelection.Item(1)

        'Retrieve all attachments from the selected item
        colAttachments = mySelectedItem.Attachments

        'Save all attachments to the selected location with a date and time stamp of message to generate a unique name
        For Each objAttachment In colAttachments
            If objAttachment.Size > 15000 Then  ' don't save attached Outlook items -- especially Notes
                MyFile = objAttachment.FileName
                DateStamp = Space(1) & Format(mySelectedItem.CreationTime, "yyyyMMddhhmmss")
                intPos = InStrRev(MyFile, ".")
                If intPos > 0 Then
                    MyFile = Left(MyFile, intPos - 1) & DateStamp & Mid(MyFile, intPos)
                Else
                    MyFile = MyFile & DateStamp
                End If
                MyFile = "C:\Scans\" & MyFile
                objAttachment.SaveAsFile(MyFile)
                intCounter = intCounter + 1
            End If
        Next
        'Cleanup
        objAttachment = Nothing
        colAttachments = Nothing
        myOlNameSpace = Nothing
        myOlSelection = Nothing
        mySelectedItem = Nothing
        If intCounter = 0 Then
            MsgBox("There are no attachments on this item larger than 15k.", vbInformation, strTitle)
        Else
            MsgBox("Saved " & intCounter & " attachment" & IIf(intCounter = 1, vbNullString, "s") & " to folder" & vbNewLine & "C:\Scans.", vbInformation, strTitle)
        End If
    End Sub

    Private Sub CopyContact2InstantFile_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles CopyContact2InstantFile.OnClick
        ' copy the active contact to InstantFile
        On Error GoTo CopyContact2InstantFile_Error
        Const strTitle As String = "Add Personal Contact to InstantFile"
        Dim olContact As Outlook.ContactItem
        Dim olNameSpace As Outlook.NameSpace
        Dim olPublicFolder As Outlook.MAPIFolder
        Dim olFolder As Outlook.MAPIFolder
        Dim olContactsFolder As Outlook.MAPIFolder
        Dim olIFContact As Outlook.ContactItem

        ' make sure a Contact is the active item
        If TypeName(OutlookApp.ActiveInspector.CurrentItem) = "ContactItem" Then
            olContact = OutlookApp.ActiveInspector.CurrentItem
            If olContact.MessageClass = "IPM.Contact.InstantFileContact" Then
                MsgBox("This already is an InstantFile Contact." & vbNewLine & "It doesn't make sense to copy it." & vbNewLine & vbNewLine & _
                            "Either" & vbNewLine & "1. [Attach] it to another matter or" & vbNewLine & vbNewLine & _
                            "2. choose [Actions], [New Contact from Same Company]" & vbNewLine & "to make a similar Contact.", vbExclamation, strTitle)
                olContact = Nothing
                Exit Sub
            End If
        Else
            MsgBox("Please display the Contact you wish to copy first," & vbNewLine & "then try this again.", vbExclamation, strTitle)
            Exit Sub
        End If

        olNameSpace = OutlookApp.GetNamespace("MAPI")
        For Each olPublicFolder In olNameSpace.Folders
            If Left(olPublicFolder.Name, Len(strPublicFolders)) = strPublicFolders Then GoTo GetContactsFolder
        Next olPublicFolder
        olNameSpace = Nothing
        MsgBox("Could not locate the 'Public Folders' folder.", vbExclamation, strTitle)
        Exit Sub

GetContactsFolder:
        olContactsFolder = Nothing
        For Each olFolder In olPublicFolder.Folders
            If olFolder.Name = "All Public Folders" Then
                For Each olContactsFolder In olFolder.Folders
                    If olContactsFolder.Name = "InstantFile Contacts" Then
                        olFolder = Nothing
                        GoTo CopyContact
                    End If
                Next olContactsFolder
            End If
        Next olFolder
        MsgBox("Could not locate the InstantFile Contacts folder.", vbExclamation, strTitle)
        If IsNothing(olContactsFolder) Then GoTo CopyContact2InstantFile_Exit

CopyContact:
        olContact.Save()  ' otherwise changes won't get written to the new contact
        olIFContact = olContactsFolder.Items.Add("IPM.Contact.InstantFileContact")
        With olIFContact
            .FullName = olContact.FullName
            .JobTitle = olContact.JobTitle
            .CompanyName = olContact.CompanyName
            .FileAs = olContact.FileAs
            .BusinessAddress = olContact.BusinessAddress
            .HomeAddress = olContact.HomeAddress
            .OtherAddress = olContact.OtherAddress
            .MailingAddress = olContact.MailingAddress
            .BusinessTelephoneNumber = olContact.BusinessTelephoneNumber
            .HomeTelephoneNumber = olContact.HomeTelephoneNumber
            .MobileTelephoneNumber = olContact.MobileTelephoneNumber
            .BusinessFaxNumber = olContact.BusinessFaxNumber
            .Email1Address = olContact.Email1Address
            .Email1AddressType = olContact.Email1AddressType
            .WebPage = olContact.WebPage
            ' from second page
            .Department = olContact.Department
            .ManagerName = olContact.ManagerName
            .AssistantName = olContact.AssistantName
            .NickName = olContact.NickName
            .Spouse = olContact.Spouse
        End With
        If MsgBox("Delete '" & olContact.Subject & "' from your personal Contacts folder?", vbQuestion + vbYesNo + vbDefaultButton2, strTitle) = vbYes Then
            olContact.Delete()
        Else
            olContact.Close(Outlook.OlInspectorClose.olSave)
        End If
        olIFContact.Display()

CopyContact2InstantFile_Exit:
        olIFContact = Nothing
        olFolder = Nothing
        olContactsFolder = Nothing
        olNameSpace = Nothing
        olContact = Nothing
        Exit Sub

CopyContact2InstantFile_Error:
        MsgBox(Err.Description, vbExclamation, strTitle)
        GoTo CopyContact2InstantFile_Exit
    End Sub

    Private Sub AdxRibbonButton2_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton2.OnClick
        ' link two open Contacts to each other
        Const strTitle As String = "Link Two Contacts to Each Other"
        Dim myInspector As Outlook.Inspector
        Dim myCont1 As Outlook.ContactItem, myCont2 As Outlook.ContactItem
        Dim strCompanyDept As String
        Dim bHave1 As Boolean
        ' make sure there are exactly two Contacts open
        myCont1 = Nothing
        For Each myInspector In OutlookApp.Inspectors
            If TypeName(myInspector.CurrentItem) = "ContactItem" Then
                If Not bHave1 Then
                    myCont1 = myInspector.CurrentItem
                    bHave1 = True
                Else
                    myCont2 = myInspector.CurrentItem
                    GoTo LinkContacts
                End If
            End If
        Next myInspector
        MsgBox("Did not find two Contacts open." & vbNewLine & vbNewLine & _
                "Open the two Contacts you want to link to each other, then try this again.", vbExclamation, strTitle)
        GoTo Link2Contacts_Exit

LinkContacts:
        ' if there are individual names in the Contacts, ask whether or not the link should display the individual's name or the company's name
        With myCont2
            strCompanyDept = .CompanyName & IIf(.Department = vbNullString, vbNullString, " (" & .Department & ")")
            If .FullName = vbNullString And .CompanyName <> vbNullString Then
                .Subject = strCompanyDept
            ElseIf .FullName <> vbNullString And .CompanyName = vbNullString Then
                .Subject = .FullName
            Else
                RetVal = MsgBox("Show the link as" & vbNewLine & "'" & .FullName & "' [Yes]" & vbNewLine & "or as" & vbNewLine & "'" & strCompanyDept & "' [No]?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Show Individual or Company Name")
                If RetVal = vbNo Then
                    .Subject = strCompanyDept
                ElseIf RetVal = vbYes Then
                    .Subject = .FullName
                ElseIf RetVal = vbCancel Then
                    GoTo Link2Contacts_Exit
                End If
            End If
            .Save()
        End With

        With myCont1
            strCompanyDept = .CompanyName & IIf(.Department = vbNullString, vbNullString, " (" & .Department & ")")
            If .FullName = vbNullString And .CompanyName <> vbNullString Then
                .Subject = strCompanyDept
            ElseIf .FullName <> vbNullString And .CompanyName = vbNullString Then
                .Subject = .FullName
            Else
                RetVal = MsgBox("Show the link as" & vbNewLine & "'" & .FullName & "' [Yes]" & vbNewLine & "or as" & vbNewLine & "'" & strCompanyDept & "' [No]?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Show Individual or Company Name")
                If RetVal = vbNo Then
                    .Subject = strCompanyDept
                ElseIf RetVal = vbYes Then
                    .Subject = .FullName
                ElseIf RetVal = vbCancel Then
                    GoTo Link2Contacts_Exit
                End If
            End If
            .Save()
        End With

        If MsgBox("LINK:" & vbNewLine & myCont1.Subject & vbNewLine & vbNewLine & _
                           "AND:" & vbNewLine & myCont2.Subject, vbQuestion + vbYesNo, strTitle) = vbYes Then
            ' link 1 to 2
            myCont1.Links.Add(myCont2)
            myCont1.Save()
            ' link 2 to 1
            myCont2.Links.Add(myCont1)
            myCont2.Save()
            MsgBox("The two Contacts were successfully linked to each other", vbInformation, strTitle)
        End If

Link2Contacts_Exit:
        myInspector = Nothing
        myCont1 = Nothing
        myCont2 = Nothing

    End Sub

    Private Sub AdxRibbonButton1_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton1.OnClick
        Const strTitle As String = "Copy Item to Drafts Folder"
        Dim olTask As Outlook.TaskItem, olNew As Outlook.TaskItem
        Dim strSubject As String, olFolder As Outlook.Folder, obj As Object, olDraft As Outlook.MailItem
        If TypeName(OutlookApp.ActiveInspector.CurrentItem) = "TaskItem" Then
            olTask = OutlookApp.ActiveInspector.CurrentItem
            ' most users don't have permission to DELETE items from NewCallTracking
            olNew = olTask.Copy()
            strSubject = olNew.Subject
            olNew.UserProperties("CallDate").Value = olTask.UserProperties("CallDate") ' otherwise olNew uses the current date/time
            ' once the item is saved, most users don't have permissions to MOVE it (deletes from NewCallTracking)
            ' if it's not saved, the MOVE fails, but without an error message
            ' olNew.Save()
            olNew.Move(OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts))
            ' if it's moved without being saved, it copies to Drafts and leaves the new item in the current folder
            olNew.UserProperties("CallerName").Value = "DELETE ME I'M A DUPLICATE"
            ' so it shows up at the top of the list, so Chuck can delete it
            olNew.UserProperties("CallDate").Value = Now
            olNew.Save()
            olNew = Nothing
            If MsgBox("The item was copied to your Drafts folder." & vbNewLine & vbNewLine & _
                      "Close the original item?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, strTitle) = vbYes Then
                olTask.Close(Outlook.OlInspectorClose.olSave)
            End If
            olTask = Nothing

            ' display the item for the user
            olFolder = OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
            For Each obj In olFolder.Items
                If TypeName(obj) = "MailItem" Then
                    olDraft = obj
                    If olDraft.Subject = strSubject Then
                        olDraft.Display()
                        Exit For
                    End If
                End If
            Next
            olDraft = Nothing
            obj = Nothing
        Else
            MsgBox("This only works with NewCallTracking or other Task type items.", vbInformation, strTitle)
        End If
    End Sub

    Private Sub DisplayMatOrDoc(ByRef myNoteItem As Outlook.NoteItem)
        On Error GoTo DisplayMatOrDoc_Error
        Const strTitle As String = "Display InstantFile Matter or Document"
        Dim appAccess As Access.Application
        Dim lngDocNo As Long, dblMatNo As Double, strID As String, intX As Integer
        Dim myInspector As Outlook.Inspector
        Dim myNotes As Outlook.Items, myNote As Outlook.NoteItem
        Dim olNameSpace As Outlook.NameSpace, olItem As Object

        If Left(myNoteItem.Subject, 18) = strIFdocNo Then
            lngDocNo = Mid(myNoteItem.Subject, 19)
            If IsDBNull(lngDocNo) Or lngDocNo = 0 Then
                MsgBox("The item does not have a DocNo", vbExclamation, "Show Document")
            Else
                appAccess = GetObject(, "Access.Application")
                With appAccess
                    If .Visible Then
                    Else
                        .Quit()
                        'appAccess = Nothing
                        Exit Sub
                    End If
                End With
                ' close and delete the Note
                For Each myInspector In OutlookApp.Inspectors
                    If Left(myInspector.Caption, 18) = strIFdocNo Then
                        myInspector.Close(Outlook.OlInspectorClose.olSave)
                        myNotes = OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderNotes).Items
                        For Each myNote In myNotes
                            If Left(myNote.Subject, 18) = strIFdocNo Then myNote.Delete()
                        Next myNote
                        myNote = Nothing
                        myNotes = Nothing
                        Exit For
                    End If
                Next myInspector
                appAccess.Run("DisplayDocument", lngDocNo)
                'appAccess = Nothing
            End If
        ElseIf Left(myNoteItem.Subject, 18) = strIFmatNo Then
            dblMatNo = Mid(myNoteItem.Subject, 19)
            If IsDBNull(dblMatNo) Or dblMatNo = 0 Then
                MsgBox("The item does not have a Matter No", vbExclamation, "Show Matter")
            Else
                ' appAccess = GetObject(, "Access.Application")
                'If Not appAccess.Visible Then appAccess.Visible = True
                'appAccess = New Access.Application
                'With appAccess
                '.Visible = True
                '.OpenCurrentDatabase("C:\Access\Access2010\GKBM\OutlookStubs.accdb", True)
                'End With
                'With appAccess
                'If .Visible Then
                'Else
                '.CloseCurrentDatabase()
                '.Quit()
                ' appAccess = Nothing
                ' End If
                'End With
                appAccess = CType(Marshal.GetActiveObject("Access.Application"), Microsoft.Office.Interop.Access.Application)
                MsgBox("appAccess was set without error")
                appAccess.Run("DisplayMatter", dblMatNo)
                appAccess = Nothing
            End If
        ElseIf Left(myNoteItem.Body, Len(strNewCallTrackingTag)) = strNewCallTrackingTag Then
            strID = Mid(myNoteItem.Body, Len(strNewCallTrackingTag) + 3)
            olNameSpace = OutlookApp.GetNamespace("MAPI")
            olItem = olNameSpace.GetItemFromID(strID, strPublicStoreID)
            olItem.Display()
        ElseIf Left(myNoteItem.Body, Len(strNewCallAppointmentTag)) = strNewCallAppointmentTag Then
            strID = Mid(myNoteItem.Body, Len(strNewCallAppointmentTag) + 3)
            olNameSpace = OutlookApp.GetNamespace("MAPI")
            olItem = olNameSpace.GetItemFromID(strID, strPublicStoreID)
            olItem.Display()
            ' added 7/23/2008 for reminders to follow up on InstantFile Requests & Tasks
        ElseIf Left(myNoteItem.Body, Len(strIFtaskTag)) = strIFtaskTag Then
            strID = Mid(myNoteItem.Body, Len(strIFtaskTag) + 3)
            intX = InStr(1, strID, vbNewLine)
            strID = Left(strID, intX - 1)
            olNameSpace = OutlookApp.GetNamespace("MAPI")
            olItem = olNameSpace.GetItemFromID(strID)  ' couldn't get this to work with the StoreID, but it works without the 2nd argument
            olItem.Display()
        End If

DisplayMatOrDoc_Exit:
        olItem = Nothing
        olNameSpace = Nothing
        Exit Sub

DisplayMatOrDoc_Error:
        If Err.Number = 429 Then
            MsgBox("Could not find the InstantFile program." & vbNewLine & vbNewLine & _
                        "Start InstantFile, then double click on the attachment again to display the item.", vbExclamation, strTitle)
        Else
            MsgBox(Err.Description, vbExclamation, strTitle)
        End If
        GoTo DisplayMatOrDoc_Exit
    End Sub

    Private Sub AdxOutlookAppEvents1_NewInspector(sender As Object, inspector As Object, folderName As String) Handles AdxOutlookAppEvents1.NewInspector
        If TypeName(inspector.CurrentItem) = "MailItem" Then
            myMailItem = inspector.CurrentItem
        ElseIf TypeName(inspector.CurrentItem) = "NoteItem" Then
            DisplayMatOrDoc(inspector.CurrentItem)
            myInsp = inspector
        End If
    End Sub

    Private Sub AdxAccessAppEvents1_OpenDatabase(sender As Object, e As EventArgs) Handles AdxAccessAppEvents1.OpenDatabase
        MsgBox("Access database is being opened", vbInformation, "AdxAccessAppEvents1_OpenDatabase")
    End Sub
End Class

