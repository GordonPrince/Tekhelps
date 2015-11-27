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


#Region " Add-in Express automatic code "

    Dim itemEvents As OutlookItemEventsClass1 = New OutlookItemEventsClass1(Me)
    Dim ItemsEvents As OutlookItemsEventsClass1 = New OutlookItemsEventsClass1(Me)

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
        ItemsEvents.ConnectTo(AddinExpress.MSO.ADXOlDefaultFolders.olFolderSentMail, True)
    End Sub

    Private Sub AddinModule_AddinBeginShutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.AddinBeginShutdown
        If ItemsEvents IsNot Nothing Then
            ItemsEvents.RemoveConnection()
            ItemsEvents = Nothing
        End If
    End Sub

#End Region

    Private Sub AdxOutlookAppEvents1_InspectorActivate(sender As Object, inspector As Object, folderName As String) Handles AdxOutlookAppEvents1.InspectorActivate
        Dim myInsp As Outlook.Inspector = Nothing
        Dim item As Object = Nothing
        Try
            myInsp = inspector
            If myInsp Is Nothing Then Return '11/25/2015 added this line
            item = myInsp.CurrentItem
            If itemEvents.IsConnected Then itemEvents.RemoveConnection()
            itemEvents.ConnectTo(item, True)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub AdxOutlookAppEvents1_ExplorerActivate(sender As Object, explorer As Object) Handles _
                AdxOutlookAppEvents1.ExplorerActivate, _
                AdxOutlookAppEvents1.ExplorerSelectionChange
        '11/25/2015 changed this to see if would prevent throwing the error that Kailey reported to me
        'Dim myExplorer As Outlook.Explorer = TryCast(explorer, Outlook.Explorer)
        'Debug.Print("AdxOutlookAppEvents1_ExplorerActivate entered")
        Dim myExplorer As Outlook.Explorer = Nothing
        Try
            myExplorer = explorer
        Catch ex As Exception
        End Try
        If myExplorer Is Nothing Then Return

        Dim sel As Outlook.Selection = Nothing
        Try
            sel = myExplorer.Selection
        Catch ex As Exception
        End Try
        If sel Is Nothing Then Return

        Dim item As Object = Nothing
        Try
            If itemEvents.IsConnected Then itemEvents.RemoveConnection()
            If sel.Count = 1 Then
                item = sel.Item(1)
                itemEvents.ConnectTo(item, True)
            End If
        Catch ex As Exception
        Finally
            If sel IsNot Nothing Then Marshal.ReleaseComObject(sel) : sel = Nothing
        End Try
    End Sub

    Private Sub AdxOutlookAppEvents1_Startup(sender As Object, e As EventArgs) Handles AdxOutlookAppEvents1.Startup
        Dim mySession As Outlook.NameSpace = Nothing
        Dim myFolder As Outlook.Folder = Nothing
        Dim myNotes As Outlook.Items = Nothing
        Dim myNote As Outlook.NoteItem = Nothing
        Dim myFolders As Outlook.Folders = Nothing
        Dim myExplorer As Outlook.Explorer = Nothing
        Dim myItems As Outlook.Items = Nothing
        Dim myItem As Object = Nothing
        Dim myUser As Outlook.Recipient = Nothing

        Try
            mySession = OutlookApp.GetNamespace("MAPI")
            myFolder = mySession.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderNotes)
            myNotes = myFolder.Items
            Dim x As Short
            For x = myNotes.Count To 1 Step -1
                myNote = myNotes(x)
                If Left(myNote.Body, 18) = strIFmatNo Or _
                    Left(myNote.Body, 18) = strIFdocNo Or _
                    Left(myNote.Body, 8) = "NewCall " Then
                    myNote.Delete()
                End If
                Marshal.ReleaseComObject(myNote)
            Next
            Marshal.ReleaseComObject(myNotes)
            Marshal.ReleaseComObject(myFolder)

            ' this won't work if the user is working offline
            ' If OutlookApp.Session.Offline Then
            If mySession.Offline Then
                MsgBox("Some InstantFile functionality will not work if you are working Offline." & vbNewLine & vbNewLine & _
                    "(To bring Outlook back Online, look in the bottom right corner of the Outlook window." & vbNewLine & _
                    "If the word 'Offline' is displayed, right-click on it, clear the checkbox to the left of 'Work offline'" & vbNewLine & _
                    "and see if you get a 'Connected' message." & vbNewLine & _
                    "If so, you've solved the problem.)", vbExclamation, "Working Offline")
            Else
                ' 11/11/2015 didn't finish doing this -- don't think it's needed anymore
                ' For Each olFolder In OutlookApp.Session.Folders
                'Dim intNote As Integer
                'Dim myFolders As Outlook.Folders = mySession.Folders
                'For x = 1 To myFolders.Count
                '    myFolder = myFolders(x)
                '    Debug.Print("myFolder.Name = " & myFolder.Name)
                '    If myFolder.Name = "Mailbox - InstantFile" Or myFolder.Name = strInstantFile Then
                '        olInstantFileInbox = olFolder.Folders("Inbox").Items
                '        olInstantFileTasks = olFolder.Folders("Tasks").Items

                '        ' delete any leftover notes from InstantFile attachments
                '        myNotes = olFolder.Folders("Notes").Items
                '        x = myNotes.Count
                '        For intNote = x To 1 Step -1
                '            myNote = myNotes(intNote)
                '            With myNote
                '                If Left(.Body, Len(strIFmatNo)) = strIFmatNo Or Left(.Body, Len(strIFdocNo)) = strIFdocNo Or Left(.Body, Len(strIFtaskTag)) = strIFtaskTag Then
                '                    ' Debug.Print .CreationTime
                '                    ' Stop
                '                    If DateDiff("h", .CreationTime, Now) > 1 Then .Delete()
                '                End If
                '            End With
                '        Next
                '        myNote = Nothing
                '        myNotes = Nothing
                'GoTo SetNewCallTracking
                '    End If
                'Next ' olFolder
                'MsgBox("Some InstantFile functions related to Tasks will not work unless you open InstantFile's Mailbox first.", vbExclamation, "InstantFile's Mailbox Not Available")

                'SetNewCallTracking:
                myFolders = mySession.Folders
                For x = 1 To myFolders.Count
                    myFolder = myFolders(x)
                    If Left(myFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
                        strPublicStoreID = myFolder.StoreID
                        Marshal.ReleaseComObject(myFolder)
                        GoTo HavePublicFolders
                    End If
                    Marshal.ReleaseComObject(myFolder)
                Next
                Marshal.ReleaseComObject(myFolders)
                MsgBox("Could not connect to Public Folders.", vbExclamation)
            End If

HavePublicFolders:
            myExplorer = OutlookApp.ActiveExplorer
            myExplorer.WindowState = Outlook.OlWindowState.olMaximized

            ' force the form to load in the user's private Tasks folder
            ' to create a new .oft file, open the form in Design mode, then SaveAs
            ' 11/11/2015 skipped this
            '        strScratch = "W:\InstantFileTask.oft"
            '        If My.Computer.FileSystem.FileExists(strScratch) Then
            '            GoTo LoadTemplate
            '        Else
            '            ' this is only used for development -- couldn't get mapping to W:\ to work 10/28/2015
            '            strScratch = "D:\W\InstantFileTask.oft"
            '            If My.Computer.FileSystem.FileExists(strScratch) Then
            'LoadTemplate:
            '                ' Dim myFD As Outlook.FormDescription
            '                objItem = OutlookApp.CreateItemFromTemplate(strScratch)
            '                objFolder = olNS.GetSharedDefaultFolder(OutlookApp.Session.CurrentUser, Outlook.OlDefaultFolders.olFolderTasks)
            '                objFD = objItem.FormDescription
            '                objFD.PublishForm(Outlook.OlFormRegistry.olFolderRegistry, objFolder)
            '            End If
            '        End If

            myUser = mySession.CurrentUser
            If myUser.Name = "Gordon Prince" Or myUser.Name = "Michael F. Montesi" Then
                Marshal.ReleaseComObject(myUser) : myUser = Nothing
                GoTo FinishedInstantFileTaskRequests
            End If
            Marshal.ReleaseComObject(myUser) : myUser = Nothing

            ' 11/22/2015 read any items so TaskRequest related emails will disappear from InstantFile's Inbox
            myFolders = mySession.Folders
            For x = 1 To myFolders.Count
                myFolder = myFolders(x)
                ' Debug.Print(myFolder.Name)
                If myFolder.Name = strInstantFile Then
                    Marshal.ReleaseComObject(myFolders)
                    myFolders = myFolder.Folders
                    Marshal.ReleaseComObject(myFolder)
                    Dim y As Short
                    For y = 1 To myFolders.Count
                        myFolder = myFolders(y)
                        If myFolder.Name = "Inbox" Then
                            myItems = myFolder.Items
                            Dim z As Short
                            For z = myItems.Count To 1 Step -1 ' as they are read they disappear from the collection
                                myItem = myItems(z)
                                ' Debug.Print("typename(myItem) = " & TypeName(myItem))
                                If TypeOf myItem Is Outlook.TaskRequestAcceptItem Or _
                                    TypeOf myItem Is Outlook.TaskRequestDeclineItem Or _
                                    TypeOf myItem Is Outlook.TaskRequestUpdateItem Then
                                    ' myTaskRequest = myItem
                                    myItem.Display()
                                    myItem.Close(Outlook.OlInspectorClose.olDiscard)
                                    ' Marshal.ReleaseComObject(myTaskRequest)
                                End If
                                Marshal.ReleaseComObject(myItem)
                            Next
                            Marshal.ReleaseComObject(myItems)
                            GoTo FinishedInstantFileTaskRequests
                        End If
                        Marshal.ReleaseComObject(myFolder)
                    Next
                    Marshal.ReleaseComObject(myFolders)
                End If
                Marshal.ReleaseComObject(myFolder)
            Next
            Marshal.ReleaseComObject(myFolders)
FinishedInstantFileTaskRequests:

        Catch ex As Exception
        Finally
            If myUser IsNot Nothing Then Marshal.ReleaseComObject(myUser) : myUser = Nothing
            If myItem IsNot Nothing Then Marshal.ReleaseComObject(myItem) : myItem = Nothing
            If myItems IsNot Nothing Then Marshal.ReleaseComObject(myItems) : myItems = Nothing
            If myExplorer IsNot Nothing Then Marshal.ReleaseComObject(myExplorer) : myExplorer = Nothing
            If myFolders IsNot Nothing Then Marshal.ReleaseComObject(myFolders) : myFolders = Nothing
            If myFolder IsNot Nothing Then Marshal.ReleaseComObject(myFolder) : myFolder = Nothing
            If myNote IsNot Nothing Then Marshal.ReleaseComObject(myNote) : myNote = Nothing
            If myNotes IsNot Nothing Then Marshal.ReleaseComObject(myNotes) : myNotes = Nothing
            If myFolder IsNot Nothing Then Marshal.ReleaseComObject(myFolder) : myFolder = Nothing
            If mySession IsNot Nothing Then Marshal.ReleaseComObject(mySession) : mySession = Nothing
        End Try
    End Sub

    Private Sub AdxOutlookAppEvents1_Quit(sender As Object, e As EventArgs) Handles AdxOutlookAppEvents1.Quit
        Dim appAccess As Access.Application = Nothing
        Dim myProject As Access.CurrentProject = Nothing
        Try
            appAccess = CType(Marshal.GetActiveObject("Access.Application"), Access.Application)
            myProject = appAccess.CurrentProject
            If Left(myProject.Name, 11) = strInstantFile Then
                MsgBox("You should close InstantFile before closing Outlook." & vbNewLine & vbNewLine & _
                        "InstantFile will now be closed, then Outlook will close.", vbExclamation)
                appAccess.Quit(Access.AcQuitOption.acQuitSaveAll)
            End If
        Catch ex As Exception
        Finally
            If myProject IsNot Nothing Then Marshal.ReleaseComObject(myProject) : myProject = Nothing
            If appAccess IsNot Nothing Then Marshal.ReleaseComObject(appAccess) : appAccess = Nothing
        End Try
    End Sub

    Private Sub AppointmentCalendar_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AppointmentCalendar.OnClick
        ActivateExplorer("Appointment Calendar")
    End Sub

    Private Sub SSIcalendar_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles SSIcalendar.OnClick
        ActivateExplorer("Appointment SSI")
    End Sub

    Private Sub NewCallTracking_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles NewCallTracking.OnClick
        ActivateExplorer("New Call Tracking")
    End Sub

    Public Sub ActivateExplorer(strFolderName As String)
        Dim myExplorer As Outlook.Explorer = Nothing
        Try
            If GetPublicFolder(strFolderName) Then
                myExplorer = OutlookApp.ActiveExplorer
                myExplorer.CurrentFolder = myPublicFolder
                Return
            Else
                MsgBox("Could not find the folder '" & strFolderName & "'", vbExclamation)
            End If
        Catch ex As Exception
            MsgBox("Could not find " & strFolderName & vbNewLine & vbNewLine & ex.Message, vbExclamation, "ActivateExplorer")
        Finally
            If myExplorer IsNot Nothing Then Marshal.ReleaseComObject(myExplorer) : myExplorer = Nothing
            If myPublicFolder IsNot Nothing Then Marshal.ReleaseComObject(myPublicFolder) : myPublicFolder = Nothing
        End Try
    End Sub

    Private Sub CopyContact2InstantFile_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles CopyContact2InstantFile.OnClick
        ' copy the active contact to InstantFile
        Const strTitle As String = "Copy Personal Contact to InstantFile"
        Dim myInsp As Outlook.Inspector = Nothing
        Dim item As Object = Nothing
        Dim olContact As Outlook.ContactItem = Nothing
        Dim olNameSpace As Outlook.NameSpace = Nothing
        Dim myFolders As Outlook.Folders = Nothing
        Dim olPublicFolder As Outlook.MAPIFolder = Nothing
        Dim olFolder As Outlook.MAPIFolder = Nothing
        Dim olContactsFolder As Outlook.MAPIFolder = Nothing
        Dim myItems As Outlook.Items = Nothing
        Dim olIFContact As Outlook.ContactItem = Nothing

        Try
            myInsp = OutlookApp.ActiveInspector
            item = myInsp.CurrentItem
            If TypeOf item Is Outlook.ContactItem Then
                ' olContact = OutlookApp.ActiveInspector.CurrentItem
                olContact = item
                If olContact.MessageClass = "IPM.Contact.InstantFileContact" Then
                    MsgBox("This already is an InstantFile Contact." & vbNewLine & _
                           "It doesn't make sense to copy it." & vbNewLine & vbNewLine & _
                            "Either" & vbNewLine & "1. [Attach] it to another matter, or" & vbNewLine & vbNewLine & _
                            "2. choose [Actions], [New Contact from Same Company]" & vbNewLine & "to make a similar Contact.", vbExclamation, strTitle)
                    Return
                End If
            Else
                MsgBox("Please display the Contact you wish to copy first," & vbNewLine & "then try this again.", vbExclamation, strTitle)
                Return
            End If
            olContact.Save()  ' otherwise changes won't get written to the new contact

            ' For Each olPublicFolder In olNameSpace.Folders
            olNameSpace = OutlookApp.GetNamespace("MAPI")
            myFolders = olNameSpace.Folders
            Dim x As Short
            For x = 1 To myFolders.Count
                olPublicFolder = myFolders(x)
                If Left(olPublicFolder.Name, Len(strPublicFolders)) = strPublicFolders Then GoTo GetContactsFolder
                Marshal.ReleaseComObject(olPublicFolder)
            Next ' olPublicFolder
            MsgBox("Could not locate the folder '" & strPublicFolders & "'.", vbExclamation, strTitle)
            Return

GetContactsFolder:
            Marshal.ReleaseComObject(myFolders)
            ' For Each olFolder In olPublicFolder.Folders
            myFolders = olPublicFolder.Folders
            For x = 1 To myFolders.Count
                olFolder = myFolders(x)
                If olFolder.Name = strAllPublicFolders Then
                    ' For Each olContactsFolder In olFolder.Folders
                    Marshal.ReleaseComObject(myFolders)
                    myFolders = olFolder.Folders
                    Dim y As Short
                    For y = 1 To myFolders.Count
                        olContactsFolder = myFolders(y)
                        If olContactsFolder.Name = "InstantFile Contacts" Then
                            GoTo CopyContact
                        End If
                        Marshal.ReleaseComObject(olContactsFolder)
                    Next ' olContactsFolder
                End If
            Next ' olFolder
            MsgBox("Could not locate the InstantFile Contacts folder.", vbExclamation, strTitle)
            Return

CopyContact:
            ' olIFContact = olContactsFolder.Items.Add("IPM.Contact.InstantFileContact")
            myItems = olContactsFolder.Items
            olIFContact = myItems.Add("IPM.Contact.InstantFileContact")
            Marshal.ReleaseComObject(myItems)
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
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, strTitle)
        Finally
            If olIFContact IsNot Nothing Then Marshal.ReleaseComObject(olIFContact) : olIFContact = Nothing
            If myItems IsNot Nothing Then Marshal.ReleaseComObject(myItems) : myItems = Nothing
            If olContactsFolder IsNot Nothing Then Marshal.ReleaseComObject(olContactsFolder) : olContactsFolder = Nothing
            If olFolder IsNot Nothing Then Marshal.ReleaseComObject(olFolder) : olFolder = Nothing
            If olPublicFolder IsNot Nothing Then Marshal.ReleaseComObject(olPublicFolder) : olPublicFolder = Nothing
            If myFolders IsNot Nothing Then Marshal.ReleaseComObject(myFolders) : myFolders = Nothing
            If olNameSpace IsNot Nothing Then Marshal.ReleaseComObject(olNameSpace) : olNameSpace = Nothing
            If olContact IsNot Nothing Then Marshal.ReleaseComObject(olContact) : olContact = Nothing
            If item IsNot Nothing Then Marshal.ReleaseComObject(item) : item = Nothing
            If myInsp IsNot Nothing Then Marshal.ReleaseComObject(myInsp) : myInsp = Nothing
        End Try
    End Sub

    Private Sub Link2Contacts2EachOther_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton2.OnClick
        ' link two open Contacts to each other
        Const strTitle As String = "Link Two Contacts to Each Other"
        Dim myInspectors As Outlook.Inspectors = Nothing
        Dim myInsp As Outlook.Inspector = Nothing
        Dim myCont1 As Outlook.ContactItem = Nothing
        Dim myCont2 As Outlook.ContactItem = Nothing
        Dim myLinks As Outlook.Links = Nothing
        Dim strCompanyDept As String

        ' make sure there are exactly two Contacts open
        Try
            ' For Each myInspector In OutlookApp.Inspectors
            myInspectors = OutlookApp.Inspectors
            For x = 1 To myInspectors.Count
                Dim bHave1 As Boolean
                myInsp = myInspectors(x)
                If TypeOf myInsp.CurrentItem Is Outlook.ContactItem Then
                    If Not bHave1 Then
                        myCont1 = myInsp.CurrentItem
                        bHave1 = True
                    Else
                        myCont2 = myInsp.CurrentItem
                        GoTo LinkContacts
                    End If
                End If
            Next 'myInspector
            MsgBox("Did not find two Contacts open." & vbNewLine & vbNewLine & _
                    "Open the two Contacts you want to link to each other, then try this again.", vbExclamation, strTitle)
            Return

LinkContacts:
            Marshal.ReleaseComObject(myInsp)
            Marshal.ReleaseComObject(myInspectors)
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
                        Return
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
                        Return
                    End If
                End If
                .Save()
            End With

            If MsgBox("LINK:" & vbNewLine & myCont1.Subject & vbNewLine & vbNewLine & _
                      "AND:" & vbNewLine & myCont2.Subject, vbQuestion + vbYesNo, strTitle) = vbYes Then
                ' link 1 to 2
                ' myCont1.Links.Add(myCont2)
                myLinks = myCont1.Links
                myLinks.Add(myCont2)
                myCont1.Save()
                Marshal.ReleaseComObject(myLinks)
                ' link 2 to 1
                ' myCont2.Links.Add(myCont1)
                myLinks = myCont2.Links
                myLinks.Add(myCont1)
                myCont2.Save()
                Marshal.ReleaseComObject(myLinks)
                MsgBox("The two Contacts were successfully linked to each other.", vbInformation, strTitle)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, strTitle)
        Finally
            If myLinks IsNot Nothing Then Marshal.ReleaseComObject(myLinks) : myLinks = Nothing
            If myCont2 IsNot Nothing Then Marshal.ReleaseComObject(myCont2) : myCont2 = Nothing
            If myCont1 IsNot Nothing Then Marshal.ReleaseComObject(myCont1) : myCont1 = Nothing
            If myInsp IsNot Nothing Then Marshal.ReleaseComObject(myInsp) : myInsp = Nothing
            If myInspectors IsNot Nothing Then Marshal.ReleaseComObject(myInspectors) : myInspectors = Nothing
        End Try
    End Sub

    Private Sub SaveAttachments_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButtonSaveAttachments.OnClick
        'TO-DO make sure it works with either an Inspector or an Explorer        

        ' copied from http://www.howto-outlook.com/howto/saveembeddedpictures.htm
        Const strTitle As String = "Save Attachments"
        Dim myInsp As Outlook.Inspector = Nothing
        Dim item As Object = Nothing
        Dim myAttachments As Outlook.Attachments = Nothing
        Dim myAttach As Outlook.Attachment = Nothing
        Dim DateStamp As String, MyFile As String
        Dim intCounter As Integer

        Try
            myInsp = OutlookApp.ActiveInspector
            item = myInsp.CurrentItem
            'Get all selected items
            ' mySelection = OutlookApp.ActiveExplorer.Selection
            ' mySelection = item.Selection
            ''Make sure at least one item is selected
            'If mySelection.Count = 0 Then
            '    RetVal = MsgBox("Please select an item first.", vbExclamation, strTitle)
            '    Exit Sub
            'End If

            ''Make sure only one item is selected
            'If mySelection.Count > 1 Then
            '    RetVal = MsgBox("Please select only one item.", vbExclamation, strTitle)
            '    Exit Sub
            'End If

            'Retrieve the selected item
            ' mySelectedItem = mySelection.Item(1)

            'Retrieve all attachments from the selected item
            ' myAttachments = mySelectedItem.Attachments
            myAttachments = item.Attachments

            'Save all attachments to the selected location with a date and time stamp of message to generate a unique name
            ' For Each myAttach In myAttachments
            Dim x As Short
            For x = 1 To myAttachments.Count
                myAttach = myAttachments(x)
                If myAttach.Size > 7000 Then  ' don't save attached Outlook items -- especially Notes
                    MyFile = myAttach.FileName
                    DateStamp = Space(1) & Format(item.CreationTime, "yyyyMMddhhmmss")
                    Dim intPos As Integer
                    intPos = InStrRev(MyFile, ".")
                    If intPos > 0 Then
                        MyFile = Left(MyFile, intPos - 1) & DateStamp & Mid(MyFile, intPos)
                    Else
                        MyFile = MyFile & DateStamp
                    End If
                    MyFile = "C:\Scans\" & MyFile
                    myAttach.SaveAsFile(MyFile)
                    intCounter = intCounter + 1
                End If
                Marshal.ReleaseComObject(myAttach)
            Next
            If intCounter = 0 Then
                MsgBox("There are no attachments on this item larger than 7k.", vbInformation, strTitle)
            Else
                MsgBox("Saved " & intCounter & " attachment" & IIf(intCounter = 1, vbNullString, "s") & " to folder" & vbNewLine & _
                       "C:\Scans.", vbInformation, strTitle)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, strTitle)

        Finally
            If myAttach IsNot Nothing Then Marshal.ReleaseComObject(myAttach) : myAttach = Nothing
            If myAttachments IsNot Nothing Then Marshal.ReleaseComObject(myAttachments) : myAttachments = Nothing
            If item IsNot Nothing Then Marshal.ReleaseComObject(item) : item = Nothing
            If myInsp IsNot Nothing Then Marshal.ReleaseComObject(myInsp) : myInsp = Nothing
        End Try
    End Sub

    Private Sub MakeAppointment_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles MakeAppointment.OnClick
        Dim myInsp As Outlook.Inspector = Nothing
        Dim item As Object = Nothing
        Dim myTask As Outlook.TaskItem = Nothing
        Dim myAttachments As Outlook.Attachments = Nothing
        Dim myAtt As Outlook.Attachment = Nothing
        Dim myProps As Outlook.UserProperties = Nothing
        Dim myPropT As Outlook.UserProperty = Nothing
        Dim myPropX As Outlook.UserProperty = Nothing
        Dim myNameSpace As Outlook.NameSpace = Nothing
        Dim myFolders As Outlook.Folders = Nothing
        Dim myFolder As Outlook.Folder = Nothing
        Dim myAllPublic As Outlook.Folder = Nothing
        Dim myApptCal As Outlook.Folder = Nothing
        Dim myPropL As Outlook.UserProperty = Nothing
        Dim myAppt As Outlook.AppointmentItem = Nothing
        Dim myItems As Outlook.Items = Nothing
        Dim mySession As Outlook.NameSpace = Nothing
        Dim myNotes As Outlook.Folder = Nothing
        Dim myNote As Outlook.NoteItem = Nothing
        Const strTitle As String = "Make New Appointment"

        Try
            myInsp = OutlookApp.ActiveInspector
        Catch
            MsgBox("Please open a " & strNewCallTrackingTag & " first.", vbInformation, strTitle)
            Return
        End Try

        Try
            item = myInsp.CurrentItem
            If TypeOf item Is Outlook.TaskItem Then
                Cursor.Current = Cursors.WaitCursor
                myTask = item
            Else
                MsgBox("This only works if a " & strNewCallTrackingTag & " is displayed.", vbInformation, strTitle)
                Return
            End If
            ' 11/14/2015 don't release item Marshal.ReleaseComObject(item)
            Marshal.ReleaseComObject(myInsp)

            myAttachments = myTask.Attachments
            If myAttachments.Count > 0 Then
                ' For Each myAtt In myAttachments
                Dim x As Short
                For x = 1 To myAttachments.Count
                    myAtt = myAttachments(x)
                    If myAtt.DisplayName = strNewCallAppointmentTag Then
                        MsgBox("This caller already has an appointment." & vbNewLine & vbNewLine & _
                            "Open the existing appointment and update it " & _
                            "(instead of making a new appointment).", vbInformation + vbOKOnly, strTitle)
                        Return
                    End If
                    Marshal.ReleaseComObject(myAtt)
                Next
            End If
            Marshal.ReleaseComObject(myAttachments)

            myProps = myTask.UserProperties
            myPropT = myProps("TypeOfCase")
            If Right(myTask.Subject, 1) = "/" Then
            Else
                myPropX = myProps("Screener")
                If Left(myPropT.Value, 2) = "SS" Then
                    myTask.Subject = myTask.Subject & "; SS; " & Left(myPropX.Value, 3) & "/"
                ElseIf Left(myPropT.Value, 1) = "A" Then
                    myTask.Subject = myTask.Subject & "; A; " & Left(myPropX.Value, 3) & "/"
                Else
                    myTask.Subject = myTask.Subject & "; " & Left(myPropT.Value, 2) & "; " & Left(myPropX.Value, 3) & "/"
                End If
                myTask.Save()
                Marshal.ReleaseComObject(myPropX)
            End If
            Marshal.ReleaseComObject(myPropT)

            myNameSpace = OutlookApp.GetNamespace("MAPI")
            myFolders = myNameSpace.Folders
            Dim s As String = Nothing
            For x = 1 To myFolders.Count
                myFolder = myFolders(x)
                s = myFolder.Name
                If Left(s, 14) = strPublicFolders Then
                    GoTo HavePublic
                End If
                Marshal.ReleaseComObject(myFolder)
            Next
            MsgBox("Could not find Outlook folder '" & strPublicFolders & "'.", vbExclamation + vbOKOnly, strTitle)
            Exit Sub
HavePublic:
            myAllPublic = myFolder.Folders(strAllPublicFolders)
            Marshal.ReleaseComObject(myFolder)
            Marshal.ReleaseComObject(myFolders)
            Marshal.ReleaseComObject(myNameSpace)

            '11/14/2015 could not get this to work without leaving an unreleased object
            myPropL = myProps("ApptLocation")
            If Len(myPropL.Value) = 0 Then
                If Left(myPropL.Value, 2) = "SS" Then
                    myPropL.Value = "SSI"
                Else
                    myPropL.Value = "Wanda"
                End If
            End If
            If myPropL.Value = "SSI" Then
                myApptCal = myAllPublic.Folders("Appointment SSI")
            Else
                myApptCal = myAllPublic.Folders("Appointment Calendar")
            End If
            Marshal.ReleaseComObject(myAllPublic)

            myItems = myApptCal.Items
            myAppt = myItems.Add
            ' 11/21/2015 don't display it until the end, since displaying fires the InspectorActivate event which releases myAppt
            ' myAppt.Display()  
            myAppt.Subject = myTask.Subject
            If myPropL.Value = "Wanda" _
                Or myPropL.Value = "219" _
                Or myPropL.Value = "SSI" Then
                myAppt.Location = myPropL.Value
            End If
            Marshal.ReleaseComObject(myPropL)

            ' add the Note with the EntryID of NewCallTracking item to the Appointment
            ' from https://www.add-in-express.com/creating-addins-blog/2012/07/16/create-outlook-task-appointment-note-email/
            myNote = TryCast(OutlookApp.CreateItem(Outlook.OlItemType.olNoteItem), Outlook.NoteItem)
            myNote.Body = strNewCallTrackingTag & Chr(13) & Chr(10) & myTask.EntryID
            myNote.Close(Outlook.OlInspectorClose.olSave)
            myAttachments = myAppt.Attachments
            myAttachments.Add(myNote, 1)
            Marshal.ReleaseComObject(myNote)
            Marshal.ReleaseComObject(myAttachments)
            myAppt.Save()

            ' add the Note with the EntryID of the Appointment to the NewCallTracking item
            If Len(myTask.Body) > 0 Then myTask.Body = myTask.Body & Chr(13) & Chr(10)
            myNote = TryCast(OutlookApp.CreateItem(Outlook.OlItemType.olNoteItem), Outlook.NoteItem)
            myNote.Body = strNewCallAppointmentTag & Chr(13) & Chr(10) & myAppt.EntryID
            myNote.Close(Outlook.OlInspectorClose.olSave)
            myAttachments = myTask.Attachments
            myAttachments.Add(myNote, 1)
            Marshal.ReleaseComObject(myNote)
            Marshal.ReleaseComObject(myAttachments)

            'can't use this -- double dot issue  : myTask.UserProperties("ApptMade").Value = "Y"
            'also this didn't release COM objects: myPropT = myTask.UserProperties("ApptMade")
            myPropT = myProps("ApptMade")
            myPropT.Value = "Y"
            Marshal.ReleaseComObject(myPropT)
            Marshal.ReleaseComObject(myProps)

            myTask.Close(Outlook.OlInspectorClose.olSave)
            Marshal.ReleaseComObject(myTask)
            myAppt.Display()
            Marshal.ReleaseComObject(myAppt)
        Catch ex As Exception
            MsgBox("An error has occured." & vbNewLine & vbNewLine & ex.Message, vbExclamation, strTitle)
        Finally
            If myNote IsNot Nothing Then Marshal.ReleaseComObject(myNote) : myNote = Nothing
            If myNotes IsNot Nothing Then Marshal.ReleaseComObject(myNotes) : myNotes = Nothing
            If mySession IsNot Nothing Then Marshal.ReleaseComObject(mySession) : mySession = Nothing
            If myItems IsNot Nothing Then Marshal.ReleaseComObject(myItems) : myItems = Nothing
            If myAppt IsNot Nothing Then Marshal.ReleaseComObject(myAppt) : myAppt = Nothing
            If myPropL IsNot Nothing Then Marshal.ReleaseComObject(myPropL) : myPropL = Nothing
            If myApptCal IsNot Nothing Then Marshal.ReleaseComObject(myApptCal) : myApptCal = Nothing
            If myAllPublic IsNot Nothing Then Marshal.ReleaseComObject(myAllPublic) : myAllPublic = Nothing
            If myFolder IsNot Nothing Then Marshal.ReleaseComObject(myFolder) : myFolder = Nothing
            If myFolders IsNot Nothing Then Marshal.ReleaseComObject(myFolders) : myFolders = Nothing
            If myNameSpace IsNot Nothing Then Marshal.ReleaseComObject(myNameSpace) : myNameSpace = Nothing
            If myPropX IsNot Nothing Then Marshal.ReleaseComObject(myPropX) : myPropX = Nothing
            If myPropT IsNot Nothing Then Marshal.ReleaseComObject(myPropT) : myPropT = Nothing
            If myProps IsNot Nothing Then Marshal.ReleaseComObject(myProps) : myProps = Nothing
            If myAtt IsNot Nothing Then Marshal.ReleaseComObject(myAtt) : myAtt = Nothing
            If myAttachments IsNot Nothing Then Marshal.ReleaseComObject(myAttachments) : myAttachments = Nothing
            If myTask IsNot Nothing Then Marshal.ReleaseComObject(myTask) : myTask = Nothing
            ' 11/14/2015 If item IsNot Nothing Then Marshal.ReleaseComObject(item) : item = Nothing
            '11/21/2015 uncommented this after not releasing item in InspectorActivate
            If item IsNot Nothing Then Marshal.ReleaseComObject(item) : item = Nothing
            If myInsp IsNot Nothing Then Marshal.ReleaseComObject(myInsp) : myInsp = Nothing
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub OpenItemFromNote_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles OpenItemFromNote.OnClick
        Const strTitle As String = "Open Item from Attached Note"
        Dim myInspectors As Outlook.Inspectors = Nothing
        Dim myInsp As Outlook.Inspector = Nothing
        Dim myExplorer As Outlook.Explorer = Nothing
        Dim mySel As Outlook.Selection = Nothing
        Dim item As Object = Nothing
        Dim myAttachments As Outlook.Attachments = Nothing
        Dim myAttach As Outlook.Attachment = Nothing

        Try
            myInspectors = OutlookApp.Inspectors
            If myInspectors.Count > 0 Then
                myInsp = OutlookApp.ActiveInspector
                item = myInsp.CurrentItem
                Debug.WriteLine("OpenItemFromNote_OnClick item = myInsp.CurrentItem")
            Else
                myExplorer = OutlookApp.ActiveExplorer
                mySel = myExplorer.Selection
                If mySel.Count = 1 Then
                    item = mySel.Item(1)
                    Debug.WriteLine("OpenItemFromNote_OnClick item = mySel.Item(1)")
                ElseIf mySel.Count = 0 Then
                    MsgBox("Please select an item before trying to open its attached Note.", vbExclamation, strTitle)
                Else
                    MsgBox("Please select only one item before trying to open its attached Note.", vbExclamation, strTitle)
                    Return
                End If
            End If

            myAttachments = item.attachments
            Dim x As Short
            For x = 1 To myAttachments.Count
                myAttach = myAttachments(x)
                If Right(myAttach.FileName, 4) = ".msg" Then
                    Debug.WriteLine("OpenItemFromNote_OnClick Right(myAttach.FileName, 4) = .msg")
                    If InterceptNote(myAttach) Then
                        Debug.WriteLine("OpenItemFromNote_OnClick InterceptNote(myAttach) returned true")
                        If myInsp IsNot Nothing Then  ' could have been triggered by an Explorer selection
                            If TypeOf item Is Outlook.AppointmentItem Or TypeOf item Is Outlook.TaskItem Then
                                myInsp.Close(Outlook.OlInspectorClose.olSave)
                            End If
                        End If
                        Return
                    End If
                End If
                Marshal.ReleaseComObject(myAttach)
            Next
            MsgBox("There are no Notes attached to this item.", vbInformation, strTitle)
            
        Catch ex As Exception
            If InStr(ex.Message, "The Explorer has been closed") Then
                MsgBox("This will not work from Outlook Today." & vbNewLine & vbNewLine & _
                       "There is no item selected.", vbExclamation, strTitle)
            Else
                MsgBox(ex.Message, vbExclamation, strTitle)
            End If
        Finally
            If myAttach IsNot Nothing Then Marshal.ReleaseComObject(myAttach) : myAttach = Nothing
            If myAttachments IsNot Nothing Then Marshal.ReleaseComObject(myAttachments) : myAttachments = Nothing
            If item IsNot Nothing Then Marshal.ReleaseComObject(item) : item = Nothing
            If mySel IsNot Nothing Then Marshal.ReleaseComObject(mySel) : mySel = Nothing
            If myExplorer IsNot Nothing Then Marshal.ReleaseComObject(myExplorer) : myExplorer = Nothing
            If myInsp IsNot Nothing Then Marshal.ReleaseComObject(myInsp) : myInsp = Nothing
            If myInspectors IsNot Nothing Then Marshal.ReleaseComObject(myInspectors) : myInspectors = Nothing
        End Try
    End Sub

    Private Sub CopyItem2DraftsFolder_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton1.OnClick
        Const strTitle As String = "E-mail Copy of This Item"
        Dim myInsp As Outlook.Inspector = Nothing
        Dim item As Object = Nothing
        Dim myTask As Outlook.TaskItem = Nothing
        Dim myProperties As Outlook.UserProperties = Nothing
        Dim myUserProp As Outlook.UserProperty = Nothing
        Dim mySession As Outlook.NameSpace = Nothing
        Dim myFolder As Outlook.Folder = Nothing
        Dim myItems As Outlook.Items = Nothing
        Dim obj As Object = Nothing
        Dim myDraft As Outlook.MailItem = Nothing
        Dim myAttachments As Outlook.Attachments = Nothing
        Dim myAttach As Outlook.Attachment = Nothing
        Dim strSubject As String = Nothing

        Try
            myInsp = OutlookApp.ActiveInspector
            item = myInsp.CurrentItem
            If TypeOf item Is Outlook.TaskItem Then
                Cursor.Current = Cursors.WaitCursor
                myTask = item
            Else
                MsgBox("This only works with NewCallTracking or other Task type items.", vbInformation, strTitle)
                Return
            End If
            Dim strFileName As String = "C:\tmp\Move.msg"
            With myTask
                .Save()
                myProperties = .UserProperties
                myUserProp = myProperties("CallerName")
                If Len(myUserProp.Value) > 0 Then
                    strSubject = myUserProp.Value
                Else
                    strSubject = .Subject
                End If
                Marshal.ReleaseComObject(myUserProp)
                Marshal.ReleaseComObject(myProperties)
                .SaveAs(strFileName)
                ' close the item that was originally opened
                .Close(Outlook.OlInspectorClose.olSave)
            End With

            ' open the copy of the item from the file in the user's Tasks folder (from where it can be moved)
            myTask = OutlookApp.CreateItemFromTemplate(strFileName)
            mySession = OutlookApp.Session
            myFolder = mySession.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
            With myTask
                .Move(myFolder)
                ' don't leave a copy of the item in the user's personal Tasks folder
                .Close(Outlook.OlInspectorClose.olDiscard)
            End With
            Marshal.ReleaseComObject(myFolder)
            Marshal.ReleaseComObject(myTask)

            ' display the new item for the user
            myFolder = mySession.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
            myItems = myFolder.Items
            Dim bHaveDraft As Boolean = False
            For x = 1 To myItems.Count
                obj = myItems(x)
                If TypeOf obj Is Outlook.MailItem Then
                    myDraft = obj
                    With myDraft
                        ' Debug.Print(".Subject=" & .Subject & ", strSubject=" & strSubject)
                        If .Subject = strSubject Then

                            ' added 11/17/2015 to strip out the Task information that isn't needed -- from Status => AnsBy:
                            Dim s As Short, strBody As String
                            strBody = .Body
                            s = InStr(strBody, "Status:")
                            If s > 0 Then
                                Dim a As Short
                                a = InStr(strBody, "AnsBy:")
                                If a > s Then
                                    strBody = Left(strBody, s - 1) & Mid(strBody, a)
                                    Debug.Print(strBody)
                                    .Body = strBody
                                End If
                            End If

                            .BCC = "NewCallTracking@gkbm.com"
                            ' delete the NCT item that's attached (as a result of the Move command)
                            myAttachments = myDraft.Attachments
                            Dim y As Short
                            For y = 1 To myAttachments.Count
                                myAttach = myAttachments(y)
                                myAttach.Delete()
                                Marshal.ReleaseComObject(myAttach)
                            Next
                            Marshal.ReleaseComObject(myAttachments)
                            .Display()
                            bHaveDraft = True
                            Exit For
                        End If
                    End With
                    Marshal.ReleaseComObject(myDraft)
                End If
                Marshal.ReleaseComObject(obj)
            Next
            If myDraft IsNot Nothing Then Marshal.ReleaseComObject(myDraft)
            Marshal.ReleaseComObject(myItems)
            Marshal.ReleaseComObject(myFolder)
            If Not bHaveDraft Then MsgBox("Could not find the new E-mail" & vbNewLine & _
                                          "Subject: " & strSubject & vbNewLine & _
                                          "in your Drafts folder.", vbExclamation, strTitle)
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, strTitle)
        Finally
            If myAttach IsNot Nothing Then Marshal.ReleaseComObject(myAttach) : myAttach = Nothing
            If myAttachments IsNot Nothing Then Marshal.ReleaseComObject(myAttachments) : myAttachments = Nothing
            If myDraft IsNot Nothing Then Marshal.ReleaseComObject(myDraft) : myDraft = Nothing
            If obj IsNot Nothing Then Marshal.ReleaseComObject(obj) : obj = Nothing
            If myItems IsNot Nothing Then Marshal.ReleaseComObject(myItems) : myItems = Nothing
            If myFolder IsNot Nothing Then Marshal.ReleaseComObject(myFolder) : myFolder = Nothing
            If mySession IsNot Nothing Then Marshal.ReleaseComObject(mySession) : mySession = Nothing
            If myUserProp IsNot Nothing Then Marshal.ReleaseComObject(myUserProp) : myUserProp = Nothing
            If myProperties IsNot Nothing Then Marshal.ReleaseComObject(myProperties) : myProperties = Nothing
            If myTask IsNot Nothing Then Marshal.ReleaseComObject(myTask) : myTask = Nothing
            If item IsNot Nothing Then Marshal.ReleaseComObject(item) : item = Nothing
            If myInsp IsNot Nothing Then Marshal.ReleaseComObject(myInsp) : myInsp = Nothing
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub CopyAttachments_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles CopyAttachments.OnClick
        Const strTitle As String = "Copy Attachments from Other Mail Items"
        Dim myInspectors As Outlook.Inspectors = Nothing
        Dim myInsp As Outlook.Inspector = Nothing
        Dim myAttachments As Outlook.Attachments = Nothing
        Dim myAttach As Outlook.Attachment = Nothing
        Dim obj As Object = Nothing
        Dim myNew As Outlook.MailItem = Nothing
        Dim myNewAttachments As Outlook.Attachments = Nothing
        Dim myOther As Outlook.MailItem = Nothing
        Dim x As Short
        Dim strFileName As String

        Try
            myInspectors = OutlookApp.Inspectors
            x = myInspectors.Count
            If x = 0 Then
                MsgBox("Display the MailItem that has the Attachments on it," & vbNewLine & _
                        "then click on this button from the new E-mail.", vbInformation, strTitle)
                Return
            End If
            myInsp = OutlookApp.ActiveInspector
            obj = myInsp.CurrentItem
            Marshal.ReleaseComObject(myInsp)
            If TypeOf obj Is Outlook.MailItem Then
                myNew = obj
                If myNew.Sent Then
                    MsgBox("This item has already been sent." & vbNewLine & vbNewLine & _
                           "Display the new E-mail and try again when it has the focus.", vbExclamation, strTitle)
                    Return
                End If
            Else
                MsgBox("This only works if the active item is the new MailItem you want to add the attachments to.", vbExclamation, strTitle)
                Return
            End If
            myNewAttachments = myNew.Attachments

            ' step through the other open items, looking for MailItems with Attachments
            Dim y As Short, intAdded As Short
            For y = x To 1 Step -1
                myInsp = myInspectors(y)
                obj = myInsp.CurrentItem
                If TypeOf obj Is Outlook.MailItem Then
                    myOther = obj
                    ' only get attachments from items that have already been sent -- including the new email will be excluded
                    If myOther.Sent Then
                        Dim z As Short
                        myAttachments = myOther.Attachments
                        z = myAttachments.Count
                        If z > 0 Then
                            ' after this question is asked the event fires that releases the COM object -- so the rest won't work
                            'RetVal = MsgBox("Copy the Attachments from the MailItem" & vbNewLine & _
                            '                "'" & myOther.Subject & "'?", vbQuestion + vbYesNoCancel, strTitle)
                            'If RetVal = vbCancel Then Return
                            'If RetVal = vbYes Then
                            z = 0
                            Dim i As Short
                            For i = 1 To myAttachments.Count
                                myAttach = myAttachments(i)
                                strFileName = "C:\tmp\" & myAttach.FileName
                                myAttach.SaveAsFile(strFileName)
                                myNewAttachments.Add(strFileName)
                                My.Computer.FileSystem.DeleteFile(strFileName)
                                z = z + 1
                                intAdded = intAdded + 1
                                Marshal.ReleaseComObject(myAttach)
                            Next ' myAttach
                            ' End If
                        End If
                        Marshal.ReleaseComObject(myAttachments)
                    End If
                    Marshal.ReleaseComObject(myOther)
                End If
                Marshal.ReleaseComObject(obj)
                Marshal.ReleaseComObject(myInsp)
            Next
            Marshal.ReleaseComObject(myNewAttachments)
            If intAdded > 0 Then
                MsgBox(IIf(intAdded = 1, "One attachment was", intAdded & " attachments were") & " added to your new item.", vbInformation, strTitle)
            Else
                MsgBox("No other MailItems with Attachments were found.", vbExclamation, strTitle)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, strTitle)
        Finally
            If myOther IsNot Nothing Then Marshal.ReleaseComObject(myOther) : myOther = Nothing
            If myNewAttachments IsNot Nothing Then Marshal.ReleaseComObject(myNewAttachments) : myNewAttachments = Nothing
            If myNew IsNot Nothing Then Marshal.ReleaseComObject(myNew) : myNew = Nothing
            If obj IsNot Nothing Then Marshal.ReleaseComObject(obj) : obj = Nothing
            If myAttach IsNot Nothing Then Marshal.ReleaseComObject(myAttach) : myAttach = Nothing
            If myAttachments IsNot Nothing Then Marshal.ReleaseComObject(myAttachments) : myAttachments = Nothing
            If myInsp IsNot Nothing Then Marshal.ReleaseComObject(myInsp) : myInsp = Nothing
            If myInspectors IsNot Nothing Then Marshal.ReleaseComObject(myInspectors) : myInspectors = Nothing
        End Try
    End Sub

    Public Function OpenItemFromID(strID As String) As Boolean
        Const strTitle As String = "OpenItemFromID()"
        If strPublicStoreID Is Nothing Then
            Debug.Print("strPublicStoreID Is Nothing")
            Stop
            MsgBox("Please call Gordon about this message:" & vbNewLine & vbNewLine & "strPublicStoreID Is Nothing", vbInformation, strTitle)
            Return False
        End If
        Dim olNameSpace As Outlook.NameSpace = Nothing
        Dim item As Object = Nothing
        Try
            olNameSpace = OutlookApp.Session
            item = olNameSpace.GetItemFromID(strID, strPublicStoreID)
            item.Display()
            Return True
        Catch
            MsgBox("The item was not found in the information store.", vbOKOnly + vbExclamation, strTitle)
            Return False
        Finally
            If item IsNot Nothing Then Marshal.ReleaseComObject(item) : item = Nothing
            If olNameSpace IsNot Nothing Then Marshal.ReleaseComObject(olNameSpace) : olNameSpace = Nothing
        End Try
    End Function

    Private Sub AboutButton_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton4.OnClick
        MsgBox("Microsoft Outlook Add-in for" & vbNewLine & _
               "Gatti, Keltner, Bienvenu & Montesi, PLC." & vbNewLine & vbNewLine & _
               "Copyright (c) 1997-2015 by Tekhelps, Inc." & vbNewLine & _
               "For further information contact Gordon Prince (901) 761-3393." & vbNewLine & vbNewLine & _
               "This version dated 2015-Nov-27  5:45.", vbInformation, "About this Add-in")
    End Sub

End Class

