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

#Region "Tekhelps definitions"

    'Public WithEvents myInspectors As Outlook.Inspectors
    'Public WithEvents myInsp As Outlook.Inspector
    'Public WithEvents myMailItem As Outlook.MailItem
    'Public WithEvents myInboxItems As Outlook.Items
    'Public WithEvents mySentItems As Outlook.Items
    'Public WithEvents myTaskItems As Outlook.Items
    'Public WithEvents olInstantFileInbox As Outlook.Items
    'Public WithEvents olInstantFileTasks As Outlook.Items
#End Region

    Private Sub ConnectToSelectedItem(ByVal selection As Outlook.Selection)
        ' 11/13/2015
        If selection IsNot Nothing Then
            If selection.Count = 1 Then
                Dim item As Object = selection.Item(1)
                If TypeOf item Is Outlook.MailItem Then
                    If itemEvents.IsConnected Then
                        itemEvents.RemoveConnection()
                    End If
                    itemEvents.ConnectTo(item, True)
                End If
                Marshal.ReleaseComObject(item) : item = Nothing
            End If
        End If
    End Sub

    'Private Sub AdxOutlookAppEvents1_ExplorerSelectionChange(sender As System.Object, explorer As System.Object) Handles AdxOutlookAppEvents1.ExplorerSelectionChange
    '    'Add-in Express forum https://www.add-in-express.com/forum/read.php?PAGEN_1=3&FID=5&TID=13430
    '    'In the same fashion you handle the ExplorerActivate event. 
    '    'That is, InspectorActivate and ExplorerActivate let you handle this scenario: 
    '    'the user presses Alt+Tab to switch between Outlook windows. 
    '    'Whenever an Outlook window becomes active, 
    '    'your code disconnects from events of the currently connected item 
    '    'and connects to events of the item which is opened (InspectorActivate) or selected (ExplorerActivate). 
    '    'The ExplorerSelectionChange allows you to follow the user selecting another item. 

    '    ' 11/10/2015 added this for CallPilot errors
    '    Dim myExplorer As Outlook.Explorer = Nothing
    '    Dim sel As Outlook.Selection = Nothing
    '    Dim outlookItem As Object = Nothing
    '    Dim myMailItem As Outlook.MailItem = Nothing
    '    Try
    '        myExplorer = CType(explorer, Outlook.Explorer)
    '        If myExplorer Is Nothing Then Return
    '        Try
    '            sel = myExplorer.Selection
    '        Catch ex As Exception
    '            'skip the exception which occurs when in certain folders such as RSS Feeds   
    '        End Try
    '        If sel Is Nothing Then Return
    '        If sel.Count = 1 Then
    '            outlookItem = sel.Item(1)
    '            If TypeOf outlookItem Is Outlook.MailItem Then
    '                myMailItem = CType(outlookItem, Outlook.MailItem)
    '                If myMailItem.Sent Then
    '                    ' disconnect from the currently connected item 
    '                    itemEvents.RemoveConnection()
    '                    ' connect to events of myMailItem  
    '                    itemEvents.ConnectTo(myMailItem, True)
    '                End If
    '            End If
    '        End If
    '    Finally
    '        If myMailItem IsNot Nothing Then Marshal.ReleaseComObject(myMailItem) : myMailItem = Nothing
    '        If outlookItem IsNot Nothing Then Marshal.ReleaseComObject(outlookItem) : outlookItem = Nothing
    '        If sel IsNot Nothing Then Marshal.ReleaseComObject(sel) : sel = Nothing
    '        ' Marshal.ReleaseComObject(myExplorer) : myExplorer = Nothing
    '    End Try
    'End Sub

    Private Sub AdxOutlookAppEvents1_ExplorerActivate(sender As Object, explorer As Object) Handles AdxOutlookAppEvents1.ExplorerActivate
        Dim theExplorer As Outlook.Explorer = Nothing
        Dim selection As Outlook.Selection = Nothing
        Try
            theExplorer = TryCast(explorer, Outlook.Explorer)
            If theExplorer IsNot Nothing Then
                ' per https://www.add-in-express.com/forum/read.php?FID=5&TID=2200
                Try
                    selection = theExplorer.Selection
                Catch
                End Try
                If selection IsNot Nothing Then
                    ConnectToSelectedItem(selection)
                End If
            End If
        Finally
            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            ' If theExplorer IsNot Nothing Then Marshal.ReleaseComObject(theExplorer) : theExplorer = Nothing
        End Try
    End Sub

    Private Sub AdxOutlookAppEvents1_InspectorActivate(sender As Object, inspector As Object, folderName As String) Handles AdxOutlookAppEvents1.InspectorActivate
        'this seems to fire only when the first Inspector window is activated, 
        'not when a second or third item is opened in another Inspector window
        'so it doesn't work for closing Notes from NewCallTracking
        ' 11/13/2015 fix this
        Dim myInsp As Outlook.Inspector = Nothing ' inspector
        Dim item As Object = Nothing ' myInsp.CurrentItem
        Dim myMailItem As Outlook.MailItem = Nothing
        'Dim mySendUsing As Object = Nothing
        Try
            myInsp = inspector
            item = myInsp.CurrentItem
            If TypeOf item Is Outlook.MailItem Then
                myMailItem = myInsp.CurrentItem
            Else
                Return
            End If
            'If myMailItem.SendUsingAccount Is Nothing Then
            'Else
            '    mySendUsing = myMailItem.SendUsingAccount
            '    If mySendUsing.DisplayName = "Microsoft Exchange" Then
            '    Else
            '        ' don't try working with CallPilot items
            '        ' MsgBox("myMailItem.SendUsingAccount.DisplayName = " & myMailItem.SendUsingAccount.DisplayName)
            '        itemEvents.RemoveConnection()
            '        Return
            '    End If
            'End If
            If myMailItem.Sent Then
                ' disconnect from the currently connected item 
                itemEvents.RemoveConnection()
                ' connect to events of myMailItem 
                itemEvents.ConnectTo(myMailItem, True)
            End If
        Catch
        Finally
            'If mySendUsing IsNot Nothing Then Marshal.ReleaseComObject(mySendUsing) : mySendUsing = Nothing

            '11/15/2015 from Andrei: Don’t release mailItem if you’ve connected to its events. 
            'And remember that item and mailitem both point to the same COM object.
            'If myMailItem IsNot Nothing Then Marshal.ReleaseComObject(myMailItem) : myMailItem = Nothing
            If item IsNot Nothing Then Marshal.ReleaseComObject(item) : item = Nothing

            'don't release myInsp -- it will release the inspector object that was passed into the procedure
        End Try
    End Sub

    Private Sub AdxOutlookAppEvents1_Startup(sender As Object, e As EventArgs) Handles AdxOutlookAppEvents1.Startup
        Dim x As Short
        'delete any leftover notes from InstantFile attachments
        'myNotes = OutlookApp.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderNotes).Items
        Dim mySession As Outlook.NameSpace = Nothing
        Dim myFolder As Outlook.Folder = Nothing
        Dim myNotes As Outlook.Items = Nothing
        Dim myNote As Outlook.NoteItem = Nothing
        Dim myFolders As Outlook.Folders = Nothing
        Dim myPublicFolder As Outlook.MAPIFolder = Nothing
        Dim myExplorer As Outlook.Explorer = Nothing

        Try
            mySession = OutlookApp.GetNamespace("MAPI")
            myFolder = mySession.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderNotes)
            myNotes = myFolder.Items
            For x = myNotes.Count To 1 Step -1
                myNote = myNotes(x)
                If Left(myNote.Body, 18) = strIFmatNo Or _
                    Left(myNote.Body, 18) = strIFdocNo Or _
                    Left(myNote.Body, 8) = "NewCall " Then
                    myNote.Delete()
                End If
            Next

            '' myInboxItems = OutlookApp.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items
            'Marshal.ReleaseComObject(myFolder)
            'myFolder = mySession.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
            'myInboxItems = myFolder.Items
            '' mySentItems = OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Items
            'Marshal.ReleaseComObject(myFolder)
            'myFolder = mySession.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail)
            'mySentItems = myFolder.Items
            '' myTaskItems = OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks).Items
            'Marshal.ReleaseComObject(myFolder)
            'myFolder = mySession.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks)
            'myTaskItems = myFolder.Items
            'Marshal.ReleaseComObject(myFolder)

            ' this won't work if the user is working offline
            ' If OutlookApp.Session.Offline Then
            If mySession.Offline Then
                MsgBox("Some InstantFile functionality will not work if you are working Offline." & vbNewLine & vbNewLine & _
                    "(To bring Outlook back Online, look in the bottom right corner of the Outlook window." & vbNewLine & _
                    "If the word 'Offline' is displayed, right-click on it, clear the checkbox to the left of 'Work offline'" & vbNewLine & _
                    "and see if you get a 'Connected' message." & vbNewLine & _
                    "If so, you've solved the problem.)", vbExclamation, "Working Offline")
            Else
                ' 11/11/2015 didn't finish doing this -- do it later 
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
                ' Dim olNS As Outlook.NameSpace, objFolder As Outlook.MAPIFolder, objItem As Outlook.TaskItem
                ' For Each myPublicFolder In OutlookApp.Session.Folders
                For x = 1 To myFolders.Count
                    myPublicFolder = myFolders(x)
                    If Left(myPublicFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
                        strPublicStoreID = myPublicFolder.StoreID
                        ' For Each olFolder In myPublicFolder.Folders
                        ' 11/11/2015 skipped this, also
                        'Dim y As Int16, myF As Outlook.Folders
                        'myF = myPublicFolder.Folders
                        'For y = 1 To myF.Count
                        '    myFolder = myF(y)
                        '    If myFolder.Name = strAllPublicFolders Then
                        '        ' For Each myNewCallTracking In olFolder.Folders
                        '        Dim n As Int16, myPF As Outlook.Folders
                        '        For n = 1 To myPF.Count
                        '            myPF = myFolder(n)
                        '            If myPF.Name = "New Call Tracking" Then 
                        GoTo HaveNewCallTracking
                        '        Next
                        '    End If
                        'Next
                    End If
                    Marshal.ReleaseComObject(myPublicFolder)
                Next
                MsgBox("You may not be able to able to view New Call Tracking items." & vbNewLine & vbNewLine & "Try to get Outlook working Online if possible.", vbExclamation, "New Call Tracking Not Available")
            End If

HaveNewCallTracking:
            ' olNS = OutlookApp.GetNamespace("MAPI")
            ' Debug.Print "ExchangeConnectionMode = " & olNS.ExchangeConnectionMode
            ' Dim intExchangeConnectionMode As Integer = olNS.ExchangeConnectionMode
            ' OutlookApp.ActiveExplorer.WindowState = Outlook.OlWindowState.olMaximized
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

        Catch ex As Exception
        Finally
            If myExplorer IsNot Nothing Then Marshal.ReleaseComObject(myExplorer) : myExplorer = Nothing
            If myFolders IsNot Nothing Then Marshal.ReleaseComObject(myFolders) : myFolders = Nothing
            If myPublicFolder IsNot Nothing Then Marshal.ReleaseComObject(myPublicFolder) : myPublicFolder = Nothing
            If myNote IsNot Nothing Then Marshal.ReleaseComObject(myNote) : myNote = Nothing
            If myNotes IsNot Nothing Then Marshal.ReleaseComObject(myNotes) : myNotes = Nothing
            If myFolder IsNot Nothing Then Marshal.ReleaseComObject(myFolder) : myFolder = Nothing
            If mySession IsNot Nothing Then Marshal.ReleaseComObject(mySession) : mySession = Nothing
        End Try
    End Sub

    Private Sub AdxOutlookAppEvents1_Quit(sender As Object, e As EventArgs) Handles AdxOutlookAppEvents1.Quit
        'Dim appAccess As Access.Application = Nothing
        'Try
        '    appAccess = CType(Marshal.GetActiveObject("Access.Application"), Access.Application)
        '    'If Left(appAccess.CurrentProject.Name, 11) = strInstantFile Then
        '    MsgBox("You should close InstantFile before closing Outlook." & vbNewLine & vbNewLine & _
        '            "InstantFile will now close, then Outlook will close.", vbCritical + vbOKOnly, "GKBM Outlook Add-in")
        '    appAccess.Quit(Access.AcQuitOption.acQuitSaveAll)
        '    'End If
        'Finally
        '    If appAccess IsNot Nothing Then Marshal.ReleaseComObject(appAccess) : appAccess = Nothing
        'End Try
    End Sub

    Private Sub AboutButton_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton4.OnClick
        MsgBox("Microsoft Outlook Add-in for" & vbNewLine & _
               "Gatti, Keltner, Bienvenu & Montesi, PLC." & vbNewLine & vbNewLine & _
               "Copyright (c) 1997-2015 by Tekhelps, Inc." & vbNewLine & _
               "For further information contact Gordon Prince (901) 761-3393." & vbNewLine & vbNewLine & _
               "This version dated 2015-Nov-13  10:45.", vbInformation, "About this Add-in")
    End Sub

    Private Sub SaveAttachments_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButtonSaveAttachments.OnClick
        ' 11/11/2015 skipped this        
        '1. release objects
        '2. make sure it works with either an Inspector or an Explorer        

        ' copied from http://www.howto-outlook.com/howto/saveembeddedpictures.htm
        Const strTitle As String = "Save Attachments"
        Dim mySelection As Outlook.Selection
        Dim mySelectedItem As Object, intPos As Integer
        Dim colAttachments As Outlook.Attachments, objAttachment As Outlook.Attachment
        Dim DateStamp As String, MyFile As String
        Dim intCounter As Integer

        'Get all selected items
        mySelection = OutlookApp.ActiveExplorer.Selection
        'Make sure at least one item is selected
        If mySelection.Count = 0 Then
            RetVal = MsgBox("Please select an item first.", vbExclamation, strTitle)
            Exit Sub
        End If

        'Make sure only one item is selected
        If mySelection.Count > 1 Then
            RetVal = MsgBox("Please select only one item.", vbExclamation, strTitle)
            Exit Sub
        End If

        'Retrieve the selected item
        mySelectedItem = mySelection.Item(1)

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
        ' 11/11/2015 skipped
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
            If olFolder.Name = strAllPublicFolders Then
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

    Private Sub Link2Contacts2EachOther_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton2.OnClick
        ' skipped
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
            MsgBox("The two Contacts were successfully linked to each other.", vbInformation, strTitle)
        End If

Link2Contacts_Exit:
    End Sub

    Private Sub CopyItem2DraftsFolder_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButton1.OnClick
        ' skipped
        Const strTitle As String = "Copy Item to Drafts Folder"
        If TypeOf OutlookApp.ActiveInspector.CurrentItem Is Outlook.TaskItem Then
            Cursor.Current = Cursors.WaitCursor
            Dim olTask As Outlook.TaskItem = OutlookApp.ActiveInspector.CurrentItem
            olTask.Save()
            Dim olNew As Outlook.TaskItem = olTask.Copy()
            Dim strSubject As String = olTask.Subject
            With olNew
                ' otherwise olNew uses the current date/time
                ' .UserProperties("CallDate").Value = olTask.UserProperties("CallDate")
                'Debug.Print("the .UserProperties aren't set in the copy")
                Dim myProp As Outlook.UserProperty
                For Each myProp In olTask.UserProperties
                    If myProp.Name = "Notes" Then
                        'ElseIf myProp.Name = "Remarks" Then
                        '    ' don't copy all the update history into the email to Wanda
                        '    .UserProperties(myProp.Name).Value = vbNullString
                    Else
                        'Try
                        '    Debug.Print("myProp." & myProp.Name & " = " & myProp.Value & _
                        '                ", .UserProperties(myProp.Name).Value = " & .UserProperties(myProp.Name).Value)
                        'Catch
                        '    Stop
                        'End Try
                        .UserProperties(myProp.Name).Value = myProp.Value
                    End If
                Next
                Try
                    ' most users don't have permissions to MOVE it (deletes from NewCallTracking)
                    .Move(OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts))
                Catch
                End Try
                ' if it didn't move (due to permissions), make these changes and then save it
                .UserProperties("Locked").Value = vbNullString
                .UserProperties("CallerName").Value = "DELETE ME I'M A DUPLICATE"
                ' purge these nightly when update NewCallTracking program runs for OLAP/Analysis
                .UserProperties("CallDate").Value = #8/8/1988#
                .Save()
            End With

            ' 11/5/2015 put this here to minimize chance of editing conflicts
            olTask.Close(Outlook.OlInspectorClose.olSave)
            'If MsgBox("The item was copied to your Drafts folder." & vbNewLine & vbNewLine & _
            '          "Close the original item?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, strTitle) = vbYes Then
            '    olTask.Close(Outlook.OlInspectorClose.olSave)
            'End If

            ' display the new item for the user
            Dim olFolder As Outlook.Folder = OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
            Dim obj As Object
            For Each obj In olFolder.Items
                If TypeOf obj Is Outlook.MailItem Then
                    Dim olDraft As Outlook.MailItem = obj
                    With olDraft
                        ' Debug.Print(".Subject = " & .Subject)
                        If .Subject = strSubject Then
                            .BCC = "NewCallTracking@gkbm.com"
                            ' delete the NCT item that's attached (as a result of the Move command)
                            Dim myAttach As Outlook.Attachment
                            For Each myAttach In olDraft.Attachments
                                myAttach.Delete()
                            Next
                            .Display()
                            Exit For
                        End If
                    End With
                End If
            Next
        Else
            MsgBox("This only works with NewCallTracking or other Task type items.", vbInformation, strTitle)
        End If
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub CopyAttachments_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles CopyAttachments.OnClick
        ' skipped
        Const strTitle As String = "Copy Attachments from Another MailItem"
        Const strMsg As String = ".msg"
        Dim myAttachment As Outlook.Attachment, strFileName As String
        Dim intX As Int16, obj As Object, myNew As Outlook.MailItem, myOther As Outlook.MailItem
        Dim intY As Int16, intZ As Int16

        intX = OutlookApp.Inspectors.Count
        If intX > 1 Then
            obj = OutlookApp.ActiveInspector.CurrentItem
            If TypeOf obj Is Outlook.MailItem Then
                myNew = obj
                If myNew.Sent Then
                    MsgBox("This item has already been sent." & vbNewLine & vbNewLine & _
                           "Display the new E-mail and try again when it has the focus.", vbExclamation, strTitle)
                    Exit Sub
                End If
            Else
                MsgBox("This only works if the active item is the new MailItem you want to add the attachments to.", vbExclamation, strTitle)
                Exit Sub
            End If
            ' step through the other open items, looking for MailItems with Attachments
            For intY = intX To 1 Step -1
                obj = OutlookApp.Inspectors(intY).CurrentItem
                If TypeOf obj Is Outlook.MailItem Then
                    myOther = obj
                    If myOther.EntryID = myNew.EntryID Then
                    Else
                        intZ = myOther.Attachments.Count
                        If intZ Then
                            RetVal = MsgBox("Copy the Attachments from the MailItem" & vbNewLine & _
                                            "'" & myOther.Subject & "'?", vbQuestion + vbYesNoCancel, strTitle)
                            If RetVal = vbCancel Then Exit Sub
                            If RetVal = vbYes Then
                                intZ = 0
                                For Each myAttachment In myOther.Attachments
                                    If Right(LCase(myAttachment.FileName), 4) = strMsg Then
                                        strFileName = "C:\tmp\" & myAttachment.FileName
                                        myAttachment.SaveAsFile(strFileName)
                                        myNew.Attachments.Add(strFileName)
                                        My.Computer.FileSystem.DeleteFile(strFileName)
                                        intZ = intZ + 1
                                    End If
                                Next myAttachment
                                MsgBox(IIf(intZ = 1, "One attachment was", intZ & " attachments were") & " added to your new item.", vbInformation, strTitle)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next
            MsgBox("No other MailItems with Attachments were found.", vbExclamation, strTitle)
        Else
            MsgBox("Display the MailItem that has the Attachments on it," & vbNewLine & "then click on this button from the new E-mail.", vbInformation, strTitle)
            Exit Sub
        End If
    End Sub

    Private Sub AdxOutlookAppEvents1_NewInspector(sender As Object, inspector As Object, folderName As String) Handles AdxOutlookAppEvents1.NewInspector
        ' 11/13/2015
        Dim item As Object = Nothing
        Dim myNote As Outlook.NoteItem = Nothing
        Try
            item = inspector.CurrentItem
            If TypeOf item Is Outlook.NoteItem Then
                myNote = item
                Dim strID As String = Nothing
                If Left(myNote.Body, Len(strNewCallTrackingTag)) = strNewCallTrackingTag Then
                    strID = Mid(myNote.Body, Len(strNewCallTrackingTag) + 3)
                ElseIf Left(myNote.Body, Len(strNewCallAppointmentTag)) = strNewCallAppointmentTag Then
                    strID = Mid(myNote.Body, Len(strNewCallAppointmentTag) + 3)
                ElseIf Left(myNote.Body, Len(strIFtaskTag)) = strIFtaskTag Then
                    strID = Mid(myNote.Body, Len(strIFtaskTag) + 3)
                End If
                If Len(strID) > 0 Then
                    If OpenItemFromID(strID) Then
                    Else
                        MsgBox("Could not open Item from ID", vbExclamation + vbOKCancel, "OpenItemFromID")
                    End If
                End If
            End If
        Catch ex As Exception
        Finally
            If myNote IsNot Nothing Then Marshal.ReleaseComObject(myNote) : myNote = Nothing
            If item IsNot Nothing Then Marshal.ReleaseComObject(item) : item = Nothing
        End Try
    End Sub

    Public Function OpenItemFromID(strID As String) As Boolean
        If strPublicStoreID Is Nothing Then
            Debug.Print("strPublicStoreID Is Nothing")
            Stop
            MsgBox("Please call Gordon about this message:" & vbNewLine & vbNewLine & "strPublicStoreID Is Nothing", vbInformation, "OpenItemFromID()")
            'Dim olPublicFolder As Outlook.Folder
            'For Each olPublicFolder In OutlookApp.Session.Folders
            '    If Left(olPublicFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
            '        ' Dim strPublicStoreID As String = olPublicFolder.StoreID
            '        For Each olFolder In olPublicFolder.Folders
            '            If olFolder.Name = strAllPublicFolders Then
            '                Dim olNameSpace As Outlook.NameSpace = OutlookApp.GetNamespace("MAPI")
            '                Try
            '                    Dim item As Object = olNameSpace.GetItemFromID(strID, strPublicStoreID)
            '                    item.Display()
            '                    Return True
            '                Catch
            '                    MsgBox("The item was not found in the information store.", vbOKOnly + vbExclamation, "OpenItemFromID()")
            '                    Return False
            '                End Try
            '            End If
            '        Next
            '    End If
            'Next
            'Return False        
        End If
        Dim olNameSpace As Outlook.NameSpace = Nothing
        Dim item As Object = Nothing
        Try
            olNameSpace = OutlookApp.Session
            item = olNameSpace.GetItemFromID(strID, strPublicStoreID)
            item.Display()
            Return True
        Catch
            MsgBox("The item was not found in the information store.", vbOKOnly + vbExclamation, "OpenItemFromID()")
            Return False
        Finally
            Marshal.ReleaseComObject(item) : item = Nothing
            Marshal.ReleaseComObject(olNameSpace) : olNameSpace = Nothing
        End Try
    End Function

    'Private Sub OpenNoteFromFile_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles OpenApptFromFile.OnClick
    '    Dim myNote As Outlook.NoteItem = OutlookApp.CreateItemFromTemplate("C:\tmp\NewCall Appointment.msg")
    '    myNote.Display()
    '    Dim myInsp As Outlook.Inspector
    '    For Each myInsp In OutlookApp.Inspectors
    '        If TypeOf myInsp.CurrentItem Is Outlook.NoteItem Then
    '            Try
    '                myInsp.Close(Outlook.OlInspectorClose.olDiscard)
    '            Catch
    '                myInsp.WindowState = Outlook.OlWindowState.olMinimized
    '            End Try
    '        End If
    '    Next
    'End Sub

    Private Sub OpenItemFromNote_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles OpenItemFromNote.OnClick
        ' look for Note attachments with the right Display property 
        ' read the EntryID from the Note
        ' open the item using the EntryID
        ' 11/11/2015 
        Dim myInsp As Outlook.Inspector = Nothing
        Dim myAttachments As Outlook.Attachments = Nothing
        Dim myTask As Outlook.TaskItem = Nothing
        Dim myAppt As Outlook.AppointmentItem = Nothing
        Dim myAttach As Outlook.Attachment = Nothing
        Dim myNote As Outlook.NoteItem = Nothing
        Try
            myInsp = OutlookApp.ActiveInspector
            Dim strOriginalType As String = TypeName(myInsp.CurrentItem)
            Dim strTitle As String = "Open Item from Attached Note"
            Dim datAppt As Date
            If TypeOf myInsp.CurrentItem Is Outlook.TaskItem Then
                myTask = myInsp.CurrentItem
                myAttachments = myTask.Attachments
            ElseIf TypeOf myInsp.CurrentItem Is Outlook.AppointmentItem Then
                myAppt = myInsp.CurrentItem
                datAppt = myAppt.Start
                myAttachments = myAppt.Attachments
            Else
                MsgBox("This only works if a NewCall Tracking or Appointment item is displayed.", vbExclamation + vbOKOnly, strTitle)
                Return
            End If
            If myAttachments.Count = 0 Then
                MsgBox("There are no Notes attached to this item.", vbInformation + vbOKOnly, strTitle)
                Return
            End If

            ' For Each myAttach In myAttachments
            Dim x As Int16
            For x = 1 To myAttachments.Count
                myAttach = myAttachments(x)
                With myAttach
                    If .FileName = strNewCallAppointmentTag & ".msg" Or .FileName = strNewCallTrackingTag & ".msg" Then
                        Dim strFileName As String = "C:\tmp\" & .FileName
                        .SaveAsFile(strFileName)
                        myNote = OutlookApp.CreateItemFromTemplate(strFileName)
                        myNote.Display()
                        ' For Each myInsp In OutlookApp.Inspectors
                        ' stepping through these backward worked, the For Each loop didn't
                        Dim y As Int16
                        For y = OutlookApp.Inspectors.Count To 1 Step -1
                            myInsp = OutlookApp.Inspectors(y)
                            ' don't close emails and other types of items -- only Appointments and Tasks and Notes
                            If TypeOf myInsp.CurrentItem Is Outlook.NoteItem Then
                                'Try
                                myInsp.Close(Outlook.OlInspectorClose.olDiscard)
                                'Catch
                                'myInsp.WindowState = Outlook.OlWindowState.olMinimized
                                'End Try
                            ElseIf TypeOf myInsp.CurrentItem Is Outlook.AppointmentItem Or TypeOf myInsp.CurrentItem Is Outlook.TaskItem Then
                                If TypeName(myInsp.CurrentItem) = strOriginalType Then
                                    myInsp.Close(Outlook.OlInspectorClose.olSave)
                                    '11/10/2015 this seems to do that same thing with the prompt in the form's VBScript
                                    'ElseIf TypeOf myInsp.CurrentItem Is Outlook.TaskItem Then
                                    '    Dim myTask As Outlook.TaskItem = myInsp.CurrentItem
                                    '    Const strField As String = "ApptDateTime"
                                    '    Try
                                    '        If myTask.UserProperties(strField).Value = datAppt Then
                                    '        Else
                                    '            myTask.UserProperties(strField).Value = datAppt
                                    '            myTask.Save()
                                    '            ' MsgBox("The Appointment date/time was changed to " & datAppt & vbNewLine &  "on the NewCallTracking item.", vbOKOnly + vbInformation, "Updated Appointment Information")
                                    '        End If
                                    '    Catch ex As Exception
                                    '    End Try
                                End If
                            End If
                        Next
                        Return
                    End If
                End With
                Marshal.ReleaseComObject(myAttach)
            Next
            MsgBox("Nothing was opened.", vbInformation + vbOKOnly, strTitle)
        Catch ex As Exception
        Finally
            If myNote IsNot Nothing Then Marshal.ReleaseComObject(myNote) : myNote = Nothing
            If myAttach IsNot Nothing Then Marshal.ReleaseComObject(myAttach) : myAttach = Nothing
            If myAppt IsNot Nothing Then Marshal.ReleaseComObject(myAppt) : myAppt = Nothing
            If myTask IsNot Nothing Then Marshal.ReleaseComObject(myTask) : myTask = Nothing
            If myAttachments IsNot Nothing Then Marshal.ReleaseComObject(myAttachments) : myAttachments = Nothing
            If myInsp IsNot Nothing Then Marshal.ReleaseComObject(myInsp) : myInsp = Nothing
        End Try
    End Sub

    Private Sub MakeAppointment_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles MakeAppointment.OnClick
        Dim myInsp As Outlook.Inspector = Nothing
        Dim item As Object = Nothing
        Dim myTask As Outlook.TaskItem = Nothing
        Dim myAttachments As Outlook.Attachments = Nothing
        Dim myAtt As Outlook.Attachment = Nothing
        'Dim myUserPropT As Outlook.UserProperty = Nothing
        'Dim myUserPropS As Outlook.UserProperty = Nothing
        Dim myNameSpace As Outlook.NameSpace = Nothing
        Dim myFolders As Outlook.Folders = Nothing
        Dim myFolder As Outlook.Folder = Nothing
        Dim myAllPublic As Outlook.Folder = Nothing
        Dim myApptCal As Outlook.Folder = Nothing
        'Dim myUserPropL As Outlook.UserProperty = Nothing
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

            'myUserPropT = myTask.UserProperties("TypeOfCase")
            'If Right(myTask.Subject, 1) = "/" Then
            'Else
            '    myUserPropS = myTask.UserProperties("Screener")
            '    If Left(myUserPropT.Value, 2) = "SS" Then
            '        myTask.Subject = myTask.Subject & "; SS; " & Left(myUserPropS.Value, 3) & "/"
            '    ElseIf Left(myUserPropT.Value, 1) = "A" Then
            '        myTask.Subject = myTask.Subject & "; A; " & Left(myUserPropS.Value, 3) & "/"
            '    Else
            '        myTask.Subject = myTask.Subject & "; " & Left(myUserPropT.Value, 2) & "; " & Left(myUserPropS.Value, 3) & "/"
            '    End If
            '    myTask.Save()
            '    Marshal.ReleaseComObject(myUserPropS)
            'End If
            'Marshal.ReleaseComObject(myUserPropT)

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
            'myUserPropL = myTask.UserProperties("ApptLocation")
            'If Len(myUserPropL.Value) = 0 Then
            '    s = myUserPropT.Value
            '    If Left(s, 2) = "SS" Then
            '        myUserPropL.Value = "SSI"
            '    Else
            '        myUserPropL.Value = "Wanda"
            '    End If
            'End If
            'If myUserPropL.Value = "SSI" Then
            '    myApptCal = myAllPublic.Folders("Appointment SSI")
            'Else
            myApptCal = myAllPublic.Folders("Appointment Calendar")
            'End If

            myItems = myApptCal.Items
            myAppt = myItems.Add
            ' 11/14/2015 don't display until everything else is done & released
            ' myAppt.Display()
            myAppt.Subject = myTask.Subject
            'If myUserPropL.Value = "Wanda" _
            '    Or myUserPropL.Value = "219" _
            '    Or myUserPropL.Value = "SSI" Then
            '    myAppt.Location = myUserPropL.Value
            'End If
            ' If myUserPropL IsNot Nothing Then Marshal.ReleaseComObject(myUserPropL)

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

            ' myTask.UserProperties("ApptMade").Value = "Y"
            'If myUserPropT IsNot Nothing Then Marshal.ReleaseComObject(myUserPropT) : myUserPropT = Nothing
            'myUserPropT = myTask.UserProperties("ApptMade")
            'myUserPropT.Value = "Y"

            myTask.Close(Outlook.OlInspectorClose.olSave)
            Marshal.ReleaseComObject(myTask)
            myAppt.Display()
            Marshal.ReleaseComObject(myAppt)
            Marshal.ReleaseComObject(myAllPublic)
        Catch ex As Exception
            MsgBox("An error has occured." & vbNewLine & vbNewLine & ex.Message, vbExclamation, strTitle)
        Finally
            If myNote IsNot Nothing Then Marshal.ReleaseComObject(myNote) : myNote = Nothing
            If myNotes IsNot Nothing Then Marshal.ReleaseComObject(myNotes) : myNotes = Nothing
            If mySession IsNot Nothing Then Marshal.ReleaseComObject(mySession) : mySession = Nothing
            If myItems IsNot Nothing Then Marshal.ReleaseComObject(myItems) : myItems = Nothing
            If myAppt IsNot Nothing Then Marshal.ReleaseComObject(myAppt) : myAppt = Nothing
            'If myUserPropL IsNot Nothing Then Marshal.ReleaseComObject(myUserPropL) : myUserPropL = Nothing
            If myApptCal IsNot Nothing Then Marshal.ReleaseComObject(myApptCal) : myApptCal = Nothing
            If myAllPublic IsNot Nothing Then Marshal.ReleaseComObject(myAllPublic) : myAllPublic = Nothing
            If myFolder IsNot Nothing Then Marshal.ReleaseComObject(myFolder) : myFolder = Nothing
            If myFolders IsNot Nothing Then Marshal.ReleaseComObject(myFolders) : myFolders = Nothing
            If myNameSpace IsNot Nothing Then Marshal.ReleaseComObject(myNameSpace) : myNameSpace = Nothing
            'If myUserPropS IsNot Nothing Then Marshal.ReleaseComObject(myUserPropS) : myUserPropS = Nothing
            'If myUserPropT IsNot Nothing Then Marshal.ReleaseComObject(myUserPropT) : myUserPropT = Nothing
            If myAtt IsNot Nothing Then Marshal.ReleaseComObject(myAtt) : myAtt = Nothing
            If myAttachments IsNot Nothing Then Marshal.ReleaseComObject(myAttachments) : myAttachments = Nothing
            If myTask IsNot Nothing Then Marshal.ReleaseComObject(myTask) : myTask = Nothing
            ' 11/14/2015 If item IsNot Nothing Then Marshal.ReleaseComObject(item) : item = Nothing
            If myInsp IsNot Nothing Then Marshal.ReleaseComObject(myInsp) : myInsp = Nothing
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub NewCallTracking_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles NewCallTracking.OnClick
        'Dim olPublicFolder As Outlook.Folder
        'For Each olPublicFolder In OutlookApp.Session.Folders
        '    If Left(olPublicFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
        '        Dim olFolder As Outlook.Folder
        '        For Each olFolder In olPublicFolder.Folders
        '            If olFolder.Name = strAllPublicFolders Then
        '                Dim olTarget As Outlook.Folder
        '                For Each olTarget In olFolder.Folders
        '                    If olTarget.Name = "New Call Tracking" Then
        '                        OutlookApp.ActiveExplorer.CurrentFolder = olTarget
        '                        Exit Sub
        '                    End If
        '                Next
        '            End If
        '        Next
        '    End If
        'Next
        'MsgBox("Could not find the folder.", vbExclamation, "Show New Call Tracking folder")
        ActivateExplorer("New Call Tracking")
    End Sub

    Private Sub AppointmentCalendar_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AppointmentCalendar.OnClick
        'Dim olPublicFolder As Outlook.Folder
        'For Each olPublicFolder In OutlookApp.Session.Folders
        '    If Left(olPublicFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
        '        Dim olFolder As Outlook.Folder
        '        For Each olFolder In olPublicFolder.Folders
        '            If olFolder.Name = strAllPublicFolders Then
        '                Dim olTarget As Outlook.Folder
        '                For Each olTarget In olFolder.Folders
        '                    If olTarget.Name = "Appointment Calendar" Then
        '                        OutlookApp.ActiveExplorer.CurrentFolder = olTarget
        '                        Exit Sub
        '                    End If
        '                Next
        '            End If
        '        Next
        '    End If
        'Next
        'MsgBox("Could not find the folder.", vbExclamation, "Show Appointment Calendar folder")
        ActivateExplorer("Appointment Calendar")
    End Sub

    Public Sub ActivateExplorer(strFolderName As String)
        Dim mySession As Outlook.NameSpace = Nothing
        Dim myFolders As Outlook.Folders = Nothing
        Dim myPublicFolder As Outlook.Folder = Nothing
        Dim myFolder As Outlook.Folder = Nothing
        Dim myTarget As Outlook.Folder = Nothing
        Dim myExplorer As Outlook.Explorer = Nothing

        Try
            Dim x As Short
            ' For Each olPublicFolder In OutlookApp.Session.Folders
            mySession = OutlookApp.Session
            myFolders = mySession.Folders
            For x = 1 To myFolders.Count
                myPublicFolder = myFolders(x)
                If Left(myPublicFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
                    Dim y As Short
                    ' For Each myFolder In myPublicFolder.Folders
                    myFolders = myPublicFolder.Folders
                    For y = 1 To myFolders.Count
                        myFolder = myFolders(y)
                        If myFolder.Name = strAllPublicFolders Then
                            myFolders = myFolder.Folders
                            Dim z As Short
                            ' For Each myTarget In olFolder.Folders
                            For z = 1 To myFolders.Count
                                myTarget = myFolders(z)
                                If myTarget.Name = strFolderName Then
                                    myExplorer = OutlookApp.ActiveExplorer
                                    myExplorer.CurrentFolder = myTarget
                                    Return
                                End If
                                Marshal.ReleaseComObject(myTarget)
                            Next
                        End If
                        Marshal.ReleaseComObject(myFolder)
                    Next
                End If
                Marshal.ReleaseComObject(myPublicFolder)
            Next
        Catch ex As Exception
        Finally
            If myExplorer IsNot Nothing Then Marshal.ReleaseComObject(myExplorer) : myExplorer = Nothing
            If myTarget IsNot Nothing Then Marshal.ReleaseComObject(myTarget) : myTarget = Nothing
            If myFolder IsNot Nothing Then Marshal.ReleaseComObject(myFolder) : myFolder = Nothing
            If myPublicFolder IsNot Nothing Then Marshal.ReleaseComObject(myPublicFolder) : myPublicFolder = Nothing
            If myFolders IsNot Nothing Then Marshal.ReleaseComObject(myFolders) : myFolders = Nothing
            If mySession IsNot Nothing Then Marshal.ReleaseComObject(mySession) : mySession = Nothing
        End Try
    End Sub
End Class

