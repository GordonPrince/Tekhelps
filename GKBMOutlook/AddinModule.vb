Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Windows.Forms
Imports AddinExpress.MSO
Imports Outlook = Microsoft.Office.Interop.Outlook

'Add-in Express Add-in Module
<GuidAttribute("7E29F01B-BDC1-47B5-B1B7-634B70EA309B"), ProgIdAttribute("GKBMOutlook.AddinModule")> _
Public Class AddinModule
    Inherits AddinExpress.MSO.ADXAddinModule

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

    Const strPublicFolders As String = "Public Folders"
 
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
        MsgBox("Outlook Add-in for" & vbNewLine & _
               "Gatti, Keltner, Bienvenu & Montesi, PLC." & vbNewLine & vbNewLine & _
               "Copyright (c) 1997-2015 by Tekhelps, Inc." & vbNewLine & _
               "For further information contact Gordon Prince (901) 761-3393." & vbNewLine & vbNewLine & _
               "This version dated 2015-Oct-21.", vbInformation, "About this Add-in")
    End Sub

    Private Sub AdxRibbonButtonSaveAttachments_OnClick(sender As Object, control As IRibbonControl, pressed As Boolean) Handles AdxRibbonButtonSaveAttachments.OnClick
        ' copied from http://www.howto-outlook.com/howto/saveembeddedpictures.htm
        Const strTitle As String = "Save Attachments"
        Dim myOlNameSpace As Outlook.NameSpace, myOlSelection As Outlook.Selection
        Dim mySelectedItem As Object, intPos As Integer
        Dim colAttachments As Outlook.Attachments, objAttachment As Outlook.Attachment
        Dim DateStamp As String, MyFile As String
        Dim intCounter As Integer
        Dim RetVal As VariantType

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
            If objAttachment.Size > 15000 Then
                MyFile = objAttachment.FileName
                DateStamp = Space(1) & Format(mySelectedItem.CreationTime, "yyyymmdd_hhnnss")
                intPos = InStrRev(MyFile, ".")
                If intPos > 0 Then
                    MyFile = Left(MyFile, intPos - 1) & DateStamp & Mid(MyFile, intPos)
                Else
                    MyFile = MyFile & DateStamp
                End If
                MyFile = "C:\Scans\" & MyFile
                objAttachment.SaveAsFile(MyFile)
                intCounter = intCounter + 1
                If intCounter = 1 Then MsgBox("Saved attachment " & IIf(intCounter > 1, "#" & intCounter & Space(1), vbNullString) & "as" & vbNewLine & MyFile, vbInformation, strTitle)
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
        ElseIf intCounter > 1 Then
            MsgBox("Saved " & intCounter - 1 & " additional attachment" & IIf(intCounter = 2, vbNullString, "s") & " to " & vbNewLine & "'C:\Scans\' folder.", vbInformation, strTitle)
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
End Class

