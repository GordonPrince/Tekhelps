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
        MsgBox("Tekhelps Add-in for GKBM Outlook" & vbNewLine & vbNewLine & "Version 2015-Oct-17", vbInformation, "About")
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
End Class

