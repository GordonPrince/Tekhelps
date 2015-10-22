Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Windows.Forms
Imports AddinExpress.MSO
Imports Outlook = Microsoft.Office.Interop.Outlook

'Add-in Express Add-in Module
<GuidAttribute("7E29F01B-BDC1-47B5-B1B7-634B70EA309B"), ProgIdAttribute("GKBMOutlook.AddinModule")> _
Public Class AddinModule
    Inherits AddinExpress.MSO.ADXAddinModule

#Region "Tekhelps definitions"
    Const strPublicFolders As String = "Public Folders"
    Dim RetVal As VariantType
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
               "This version dated 2015-Oct-22 10:35.", vbInformation, "About this Add-in")
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
        Const strTitle As String = "Copy Contact to InstantFile Contacts"
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
        If TypeName(OutlookApp.ActiveInspector.CurrentItem) = "TaskItem" Then
            olTask = OutlookApp.ActiveInspector.CurrentItem
            ' there must be something different about copying / moving a TaskItem to an E-mail folder (drafts)
            olNew = olTask.Copy()
            olNew.UserProperties("CallDate").Value = olTask.UserProperties("CallDate") ' otherwise olNew uses the current date/time
            olNew.Save()
            olNew.Move(OutlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts))
            olNew = Nothing
            olTask = Nothing
            MsgBox("The item was copied to your Drafts folder.", vbInformation, strTitle)
        Else
            MsgBox("This only works with NewCallTracking or other Task type items.", vbInformation, strTitle)
        End If
    End Sub
End Class

