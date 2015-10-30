Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Windows.Forms
Imports AddinExpress.MSO
Imports System.Diagnostics
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Access = Microsoft.Office.Interop.Access

'Add-in Express Add-in Module
<GuidAttribute("7E29F01B-BDC1-47B5-B1B7-634B70EA309B"), ProgIdAttribute("GKBMOutlook.AddinModule")> _
Public Class AddinModule
    Inherits AddinExpress.MSO.ADXAddinModule

#Region "Tekhelps definitions"
    Const strPublicFolders As String = "Public Folders"
    Const strInstantFile As String = "InstantFile"
    Const strIFmatNo As String = "InstantFile_MatNo_"
    Const strIFdocNo As String = "InstantFile_DocNo_"
    Const strNewCallTrackingTag As String = "NewCall Tracking Item"
    Const strIFtaskTag As String = "InstantFile_Task"
    Const strNewCallAppointmentTag As String = "NewCall Appointment"
    Const strConnectionString As String = "App=OutlookVBA;Provider=MSDataShape.1;Persist Security Info=False;Data Source=SQLserver;Integrated Security=SSPI;" & _
                                                         "Initial Catalog=InstantFile;Data Provider=SQLOLEDB.1"
    Public strPublicStoreID As String
    Public WithEvents myInspectors As Outlook.Inspectors
    Public WithEvents myInsp As Outlook.Inspector
    Public WithEvents myMailItem As Outlook.MailItem
    Public WithEvents myInboxItems As Outlook.Items
    Public WithEvents mySentItems As Outlook.Items
    Public WithEvents myTaskItems As Outlook.Items
    Public WithEvents olInstantFileInbox As Outlook.Items
    Public WithEvents olInstantFileTasks As Outlook.Items
    Dim RetVal As VariantType
    Dim strScratch As String, lngX As Long
    Dim intExchangeConnectionMode As Integer
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

    Private itemEvents As OutlookItemEventsClass1 = Nothing

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

    Private Sub AddinModule_AddinStartupComplete(sender As System.Object, e As System.EventArgs) Handles MyBase.AddinStartupComplete
        itemEvents = New OutlookItemEventsClass1(Me)
    End Sub

#End Region

    Private Sub AdxRibbonButton4_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton4.OnClick
        MsgBox("Microsoft Outlook Add-in for" & vbNewLine & _
               "Gatti, Keltner, Bienvenu & Montesi, PLC." & vbNewLine & vbNewLine & _
               "Copyright (c) 1997-2015 by Tekhelps, Inc." & vbNewLine & _
               "For further information contact Gordon Prince (901) 761-3393." & vbNewLine & vbNewLine & _
               "This version dated 2015-Oct-30  11:55.", vbInformation, "About this Add-in")
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
            If objAttachment.Size > 7000 Then  ' don't save attached Outlook items -- especially Notes
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
        If intCounter = 0 Then
            MsgBox("There are no attachments on this item larger than 7k.", vbInformation, strTitle)
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
        If TypeOf OutlookApp.ActiveInspector.CurrentItem Is Outlook.ContactItem Then
            olContact = OutlookApp.ActiveInspector.CurrentItem
            If olContact.MessageClass = "IPM.Contact.InstantFileContact" Then
                MsgBox("This already is an InstantFile Contact." & vbNewLine & "It doesn't make sense to copy it." & vbNewLine & vbNewLine & _
                            "Either" & vbNewLine & "1. [Attach] it to another matter or" & vbNewLine & vbNewLine & _
                            "2. choose [Actions], [New Contact from Same Company]" & vbNewLine & "to make a similar Contact.", vbExclamation, strTitle)
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
        MsgBox("Could not locate the 'Public Folders' folder.", vbExclamation, strTitle)
        Exit Sub

GetContactsFolder:
        olContactsFolder = Nothing
        For Each olFolder In olPublicFolder.Folders
            If olFolder.Name = "All Public Folders" Then
                For Each olContactsFolder In olFolder.Folders
                    If olContactsFolder.Name = "InstantFile Contacts" Then
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
            If TypeOf myInspector.CurrentItem Is Outlook.ContactItem Then
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
    End Sub

    Private Sub AdxRibbonButton1_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton1.OnClick
        Const strTitle As String = "Copy Item to Drafts Folder"
        Dim olTask As Outlook.TaskItem, olNew As Outlook.TaskItem
        Dim strSubject As String, olFolder As Outlook.Folder, obj As Object, olDraft As Outlook.MailItem
        If TypeOf OutlookApp.ActiveInspector.CurrentItem Is Outlook.TaskItem Then
            olTask = OutlookApp.ActiveInspector.CurrentItem
            ' most users don't have permission to DELETE items from NewCallTracking
            olNew = olTask.Copy()
            With olNew
                strSubject = .Subject
                ' otherwise olNew uses the current date/time
                .UserProperties("CallDate").Value = olTask.UserProperties("CallDate")
                ' so opening the item doesn't prompt with the Locked by user message
                .UserProperties("Locked").Value = vbNullString

                ' once the item is saved, most users don't have permissions to MOVE it (deletes from NewCallTracking)
                ' if it's not saved, the MOVE fails, but without an error message
                .Save()
                .Move(OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts))
                ' if it's moved without being saved, it copies to Drafts and leaves the new item in the current folder
                .UserProperties("CallerName").Value = "DELETE ME I'M A DUPLICATE"
                ' purge these automatically somehow
                .UserProperties("CallDate").Value = #8/8/1988#
                .Save()
            End With
            If MsgBox("The item was copied to your Drafts folder." & vbNewLine & vbNewLine & _
                      "Close the original item?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, strTitle) = vbYes Then
                olTask.Close(Outlook.OlInspectorClose.olSave)
            End If

            ' display the item for the user
            olFolder = OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
            For Each obj In olFolder.Items
                If TypeOf obj Is Outlook.MailItem Then
                    olDraft = obj
                    If olDraft.Subject = strSubject Then
                        olDraft.Display()
                        Exit For
                    End If
                End If
            Next
        Else
            MsgBox("This only works with NewCallTracking or other Task type items.", vbInformation, strTitle)
        End If
    End Sub

    '    Private Sub AdxOutlookAppEvents1_NewInspector(sender As Object, inspector As Object, folderName As String) Handles AdxOutlookAppEvents1.NewInspector
    '        On Error GoTo DisplayMatOrDoc_Error
    '        Const strTitle As String = "Display InstantFile Matter or Document"
    '        Dim appAccess As Access.Application
    '        Dim lngDocNo As Long, dblMatNo As Double, strID As String, intX As Integer
    '        Dim myInspector As Outlook.Inspector
    '        Dim myNotes As Outlook.Items, myNote As Outlook.NoteItem
    '        Dim olNameSpace As Outlook.NameSpace, olItem As Object

    '        If TypeOf inspector.CurrentItem Is Outlook.MailItem Then
    '            myMailItem = inspector.CurrentItem
    '        ElseIf TypeOf inspector.CurrentItem Is Outlook.NoteItem Then
    '            myNote = inspector.CurrentItem
    '            ' Note: connecting to Access only works if Access and VS are running as the same user
    '            ' Especially, if Visual Studio is running as Administrator (e.g., for creating Add-ins), 
    '            ' Access must also be running as Administrator

    '            If Left(myNote.Subject, 18) = strIFdocNo Then
    '                lngDocNo = Mid(myNote.Subject, 19)
    '                If IsDBNull(lngDocNo) Or lngDocNo = 0 Then
    '                    MsgBox("The item does not have a DocNo.", vbExclamation, "Show Document")
    '                Else
    '                    appAccess = CType(Marshal.GetActiveObject("Access.Application"), Microsoft.Office.Interop.Access.Application)
    '                    appAccess.Run("DisplayDocument", lngDocNo)
    '                End If
    '            ElseIf Left(myNote.Subject, 18) = strIFmatNo Then
    '                dblMatNo = Mid(myNote.Subject, 19)
    '                If IsDBNull(dblMatNo) Or dblMatNo = 0 Then
    '                    MsgBox("The item does not have a MatterNo.", vbExclamation, "Show Matter")
    '                Else
    '                    appAccess = CType(Marshal.GetActiveObject("Access.Application"), Microsoft.Office.Interop.Access.Application)
    '                    appAccess.Run("DisplayMatter", dblMatNo)
    '                End If
    '            ElseIf Left(myNote.Body, Len(strNewCallTrackingTag)) = strNewCallTrackingTag Then
    '                strID = Mid(myNote.Body, Len(strNewCallTrackingTag) + 3)
    '                olNameSpace = OutlookApp.GetNamespace("MAPI")
    '                olItem = olNameSpace.GetItemFromID(strID, strPublicStoreID)
    '                olItem.Display()
    '            ElseIf Left(myNote.Body, Len(strNewCallAppointmentTag)) = strNewCallAppointmentTag Then
    '                strID = Mid(myNote.Body, Len(strNewCallAppointmentTag) + 3)
    '                olNameSpace = OutlookApp.GetNamespace("MAPI")
    '                olItem = olNameSpace.GetItemFromID(strID, strPublicStoreID)
    '                olItem.Display()
    '            ElseIf Left(myNote.Body, Len(strIFtaskTag)) = strIFtaskTag Then
    '                strID = Mid(myNote.Body, Len(strIFtaskTag) + 3)
    '                intX = InStr(1, strID, vbNewLine)
    '                strID = Left(strID, intX - 1)
    '                olNameSpace = OutlookApp.GetNamespace("MAPI")
    '                olItem = olNameSpace.GetItemFromID(strID)  ' couldn't get this to work with the StoreID, but it works without the 2nd argument
    '                olItem.Display()
    '            End If
    '        End If
    '        Exit Sub

    'DisplayMatOrDoc_Error:
    '        If Err.Number = 429 Then
    '            MsgBox("Could not find the InstantFile program." & vbNewLine & vbNewLine & _
    '                    "Start InstantFile, then double click on the attachment again to display the item.", vbExclamation, strTitle)
    '        Else
    '            MsgBox(Err.Description, vbExclamation, strTitle)
    '        End If
    '    End Sub

    Private Sub AdxOutlookAppEvents1_Startup(sender As Object, e As EventArgs) Handles AdxOutlookAppEvents1.Startup
        On Error GoTo Startup_Error
        Const strTitle As String = "AdxOutlookAppEvents1_Startup()"
        Dim intX As Integer
        Dim olPublicFolder As Outlook.MAPIFolder, olFolder As Outlook.MAPIFolder
        Dim olNS As Outlook.NameSpace, objFolder As Outlook.MAPIFolder, objItem As Outlook.TaskItem
        Dim objFD As Outlook.FormDescription
        Dim intHour As Integer
        Dim intNote As Integer, myNotes As Outlook.Items, myNote As Outlook.NoteItem
        Dim olRem As Outlook.Reminder

        ' delete any leftover notes from InstantFile attachments
        myNotes = OutlookApp.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderNotes).Items
        intX = myNotes.Count
        For intNote = intX To 1 Step -1
            myNote = myNotes(intNote)
            If Left(myNote.Body, 18) = strIFmatNo Or _
                Left(myNote.Body, 18) = strIFdocNo Or _
                Left(myNote.Body, 8) = "NewCall " Then _
                myNote.Delete()
        Next

        ' Set myInspectors = Application.Inspectors
        myInboxItems = OutlookApp.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items
        mySentItems = OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Items
        myTaskItems = OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks).Items

        ' this won't work if the user is working offline
        If OutlookApp.Session.Offline Then
            MsgBox("Some InstantFile functionality will not work if you are working Offline." & vbNewLine & vbNewLine & _
                "(To bring Outlook back Online, look in the bottom right corner of the Outlook window." & vbNewLine & _
                "If the word 'Offline' is displayed, right-click on it, clear the checkbox to the left of 'Work offline'" & vbNewLine & _
                "and see if you get a 'Connected' message." & vbNewLine & _
                "If so, you've solved the problem.)", vbExclamation, "Working Offline")
        Else
            For Each olFolder In OutlookApp.Session.Folders
                ' Debug.Print olFolder.Name
                If olFolder.Name = "Mailbox - InstantFile" Or olFolder.Name = strInstantFile Then
                    olInstantFileInbox = olFolder.Folders("Inbox").Items
                    olInstantFileTasks = olFolder.Folders("Tasks").Items

                    ' delete any leftover notes from InstantFile attachments
                    myNotes = olFolder.Folders("Notes").Items
                    intX = myNotes.Count
                    For intNote = intX To 1 Step -1
                        myNote = myNotes(intNote)
                        With myNote
                            If Left(.Body, Len(strIFmatNo)) = strIFmatNo Or Left(.Body, Len(strIFdocNo)) = strIFdocNo Or Left(.Body, Len(strIFtaskTag)) = strIFtaskTag Then
                                ' Debug.Print .CreationTime
                                ' Stop
                                If DateDiff("h", .CreationTime, Now) > 1 Then .Delete()
                            End If
                        End With
                    Next
                    myNote = Nothing
                    myNotes = Nothing
                    GoTo SetNewCallTracking
                End If
            Next olFolder
            MsgBox("Some InstantFile functions related to Tasks will not work unless you open InstantFile's Mailbox first.", vbExclamation, "InstantFile's Mailbox Not Available")

SetNewCallTracking:
            For Each olPublicFolder In OutlookApp.Session.Folders
                If Left(olPublicFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
                    strPublicStoreID = olPublicFolder.StoreID
                    For Each olFolder In olPublicFolder.Folders
                        If olFolder.Name = "All Public Folders" Then
                            For Each myNewCallTracking In olFolder.Folders
                                If myNewCallTracking.Name = "New Call Tracking" Then GoTo HaveNewCallTracking
                            Next
                        End If
                    Next
                End If
            Next
            MsgBox("You may not be able to able to view New Call Tracking items." & vbNewLine & vbNewLine & "Try to get Outlook working Online if possible.", vbExclamation, "New Call Tracking Not Available")
        End If

HaveNewCallTracking:
        olNS = OutlookApp.GetNamespace("MAPI")
        ' Debug.Print "ExchangeConnectionMode = " & olNS.ExchangeConnectionMode
        intExchangeConnectionMode = olNS.ExchangeConnectionMode
        OutlookApp.ActiveExplorer.WindowState = Outlook.OlWindowState.olMaximized
        ' force the form to load in the user's private Tasks folder
        ' to create a new .oft file, open the form in Design mode, then SaveAs
        strScratch = "W:\InstantFileTask.oft"
        If My.Computer.FileSystem.FileExists(strScratch) Then
            GoTo LoadTemplate
        Else
            ' this is only used for development -- couldn't get mapping to W:\ to work 10/28/2015
            strScratch = "D:\W\InstantFileTask.oft"
            If My.Computer.FileSystem.FileExists(strScratch) Then
LoadTemplate:
                objItem = OutlookApp.CreateItemFromTemplate(strScratch)
                objFolder = olNS.GetSharedDefaultFolder(OutlookApp.Session.CurrentUser, Outlook.OlDefaultFolders.olFolderTasks)
                objFD = objItem.FormDescription
                objFD.PublishForm(Outlook.OlFormRegistry.olFolderRegistry, objFolder)
            End If
        End If
        Exit Sub

Startup_Error:
        MsgBox(Err.Description, vbExclamation, strTitle)
    End Sub

    Private Sub AdxOutlookAppEvents1_Quit(sender As Object, e As EventArgs) Handles AdxOutlookAppEvents1.Quit
        On Error GoTo AdxOutlookAppEvents1_Error
        Dim appAccess As Access.Application
        appAccess = CType(Marshal.GetActiveObject("Access.Application"), Microsoft.Office.Interop.Access.Application)
        ' If appAccess.CurrentProject.Name = "OutlookStubs.accdb" Then
        If Left(appAccess.CurrentProject.Name, 11) = strInstantFile Then
            MsgBox("InstantFile should be closed before Outlook is closed." & vbNewLine & vbNewLine & _
                    "InstantFile will now close, then Outlook will close.", vbCritical + vbOKOnly, "Warning")
            appAccess.Quit(Access.AcQuitOption.acQuitSaveNone)
        End If
        Exit Sub

AdxOutlookAppEvents1_Error:
        If Err.Number = 429 Or Err.Number = 2467 Then
            ' Access not running
        Else
            MsgBox(Err.Description, vbExclamation, "AdxOutlookAppEvents1_Quit")
        End If
    End Sub

    Private Sub ConnectToSelectedItem(ByVal selection As Outlook.Selection)
        If selection IsNot Nothing Then
            If selection.Count = 1 Then
                Dim item As Object = selection.Item(1)
                If TypeOf item Is Outlook.MailItem Then
                    If itemEvents.IsConnected Then
                        itemEvents.RemoveConnection()
                        Debug.Print("Disconnected from the previously connected item.")
                    End If
                    itemEvents.ConnectTo(item, True)
                    Debug.Print("Connected to this Outlook item.")
                Else
                    Marshal.ReleaseComObject(item)
                    Debug.Print("Do not connect to this Outlook item.")
                End If
            End If
        End If
    End Sub

    Private Sub AdxOutlookAppEvents1_ExplorerActivate(sender As Object, explorer As Object) Handles AdxOutlookAppEvents1.ExplorerActivate
        Dim theExplorer As Outlook.Explorer = TryCast(explorer, Outlook.Explorer)
        If theExplorer IsNot Nothing Then
            Dim selection As Outlook.Selection = Nothing
            Try
                selection = theExplorer.Selection
            Catch
            End Try

            If selection IsNot Nothing Then
                ConnectToSelectedItem(selection)
                Marshal.ReleaseComObject(selection)
            End If
        End If
    End Sub

    'Private Sub AdxOutlookAppEvents1_NewInspector(sender As Object, inspector As Object, folderName As String) Handles AdxOutlookAppEvents1.NewInspector
    '    MsgBox("The AdxOutlookAppEvents1_NewInspector() event has occurred")
    'End Sub

    Private Sub AdxOutlookAppEvents1_InspectorActivate(sender As Object, inspector As Object, folderName As String) Handles AdxOutlookAppEvents1.InspectorActivate
        Dim myInsp As Outlook.Inspector, myMailItem As Outlook.MailItem
        myInsp = inspector
        'Debug.Print("Entered AdxOutlookAppEvents1_InspectorActivate() at " & Now)
        If TypeOf myInsp.CurrentItem Is Outlook.MailItem Then
            myMailItem = myInsp.CurrentItem
            If myMailItem.Sent Then
                'Debug.WriteLine("myMailItem.Sent = " & myMailItem.Sent)
                'MsgBox("myMailItem.Sent = " & myMailItem.Sent)
                'Debug.WriteLine("after MsgBox(myMailItem.Sent = " & myMailItem.Sent & ")")
                Dim theInspector As Outlook.Inspector = TryCast(inspector, Outlook.Inspector)

                If theInspector IsNot Nothing Then
                    'Debug.Print("theInspector IsNot Nothing")
                    Dim selection As Outlook.Selection = Nothing
                    Try
                        selection = theInspector.Application.ActiveExplorer.Selection
                    Catch
                    End Try

                    If selection IsNot Nothing Then
                        'Debug.Print("selection IsNot Nothing")
                        ConnectToSelectedItem(selection)
                        'Debug.Print("ConnectToSelectedItem(selection) finished")
                        Marshal.ReleaseComObject(selection)
                    Else
                        'Debug.Print("selection is Nothing")
                    End If
                Else
                    'Debug.Print("theInspector Is Nothing")
                End If
            End If
        End If
        'Debug.Print("Exiting AdxOutlookAppEvents1_InspectorActivate()")
    End Sub
End Class

