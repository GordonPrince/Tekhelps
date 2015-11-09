﻿Imports System.Windows.Forms

Partial Public Class AddinModule

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New()

        Application.EnableVisualStyles()

        'This call is required by the Component Designer
        InitializeComponent()

        'Please add any initialization code to the AddinInitialize event handler

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.AdxRibbonTab1 = New AddinExpress.MSO.ADXRibbonTab(Me.components)
        Me.AdxRibbonGroup1 = New AddinExpress.MSO.ADXRibbonGroup(Me.components)
        Me.SaveClose = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonSeparator1 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        Me.CopyContact2InstantFile = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonButton2 = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonButtonSaveAttachments = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.MakeAppointment = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.OpenItemFromNote = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonSeparator2 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        Me.AdxRibbonButton1 = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.CopyAttachments = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonSeparator3 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        Me.AdxRibbonButton4 = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.OpenApptFromFile = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxOutlookAppEvents1 = New AddinExpress.MSO.ADXOutlookAppEvents(Me.components)
        Me.AdxRibbonSeparator4 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        Me.AdxRibbonSeparator5 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        '
        'AdxRibbonTab1
        '
        Me.AdxRibbonTab1.Caption = "GKBM"
        Me.AdxRibbonTab1.Controls.Add(Me.AdxRibbonGroup1)
        Me.AdxRibbonTab1.Id = "adxRibbonTab_a5cf84a1ced7485ea57b9386db6cf9ed"
        Me.AdxRibbonTab1.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
        '
        'AdxRibbonGroup1
        '
        Me.AdxRibbonGroup1.Caption = "Custom Functions"
        Me.AdxRibbonGroup1.Controls.Add(Me.SaveClose)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonSeparator1)
        Me.AdxRibbonGroup1.Controls.Add(Me.CopyContact2InstantFile)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonSeparator5)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButton2)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButtonSaveAttachments)
        Me.AdxRibbonGroup1.Controls.Add(Me.MakeAppointment)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonSeparator4)
        Me.AdxRibbonGroup1.Controls.Add(Me.OpenItemFromNote)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonSeparator2)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButton1)
        Me.AdxRibbonGroup1.Controls.Add(Me.CopyAttachments)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonSeparator3)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButton4)
        Me.AdxRibbonGroup1.Controls.Add(Me.OpenApptFromFile)
        Me.AdxRibbonGroup1.Id = "adxRibbonGroup_072621b7e27f4dc6966496c4216ab446"
        Me.AdxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AdxRibbonGroup1.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
        '
        'SaveClose
        '
        Me.SaveClose.Caption = "Save && Close"
        Me.SaveClose.Id = "adxRibbonButton_3d19a766ad2c4d30be39ba838a25052e"
        Me.SaveClose.IdMso = "SaveAndClose"
        Me.SaveClose.ImageMso = "SaveAndClose"
        Me.SaveClose.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.SaveClose.Ribbons = CType(((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose Or AddinExpress.MSO.ADXRibbons.msrOutlookMeetingRequestRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookPostRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSRead), AddinExpress.MSO.ADXRibbons)
        Me.SaveClose.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonSeparator1
        '
        Me.AdxRibbonSeparator1.Id = "adxRibbonSeparator_644376ddea174afdbe6095b070a6339c"
        Me.AdxRibbonSeparator1.Ribbons = CType(((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose Or AddinExpress.MSO.ADXRibbons.msrOutlookMeetingRequestRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookPostRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSRead), AddinExpress.MSO.ADXRibbons)
        '
        'CopyContact2InstantFile
        '
        Me.CopyContact2InstantFile.Caption = "Copy Contact to InstantFile"
        Me.CopyContact2InstantFile.Id = "adxRibbonButton_0e8a15c449ee4798adf475a1587fd047"
        Me.CopyContact2InstantFile.ImageMso = "RecordsAddFromOutlook"
        Me.CopyContact2InstantFile.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.CopyContact2InstantFile.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookContact
        Me.CopyContact2InstantFile.ScreenTip = "Copy this personal Contact to InstantFile's Contacts"
        Me.CopyContact2InstantFile.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonButton2
        '
        Me.AdxRibbonButton2.Caption = "Link Two Contacts"
        Me.AdxRibbonButton2.Id = "adxRibbonButton_4b812ad786fc4beb9f6a39f3bbcb0352"
        Me.AdxRibbonButton2.ImageMso = "ObjectsGroup"
        Me.AdxRibbonButton2.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AdxRibbonButton2.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookContact
        Me.AdxRibbonButton2.ScreenTip = "Link two Contacts to each other."
        Me.AdxRibbonButton2.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonButtonSaveAttachments
        '
        Me.AdxRibbonButtonSaveAttachments.Caption = "Save Attachments"
        Me.AdxRibbonButtonSaveAttachments.Id = "adxRibbonButton_4c9aad6ef1324d11bf940cb9a7ca7623"
        Me.AdxRibbonButtonSaveAttachments.ImageMso = "SaveAttachments"
        Me.AdxRibbonButtonSaveAttachments.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AdxRibbonButtonSaveAttachments.Ribbons = CType(((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookPostRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSRead), AddinExpress.MSO.ADXRibbons)
        Me.AdxRibbonButtonSaveAttachments.ScreenTip = "Save Attachments to C:\Scans folder."
        Me.AdxRibbonButtonSaveAttachments.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'MakeAppointment
        '
        Me.MakeAppointment.Caption = "Make Appointment"
        Me.MakeAppointment.Id = "adxRibbonButton_6125721066004d2a94e0fb7b2a58af6f"
        Me.MakeAppointment.ImageMso = "NewAppointment"
        Me.MakeAppointment.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.MakeAppointment.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookTask
        Me.MakeAppointment.ScreenTip = "Make an Appointment for this NewCallTracking item."
        Me.MakeAppointment.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'OpenItemFromNote
        '
        Me.OpenItemFromNote.Caption = "Open Item from Note"
        Me.OpenItemFromNote.Id = "adxRibbonButton_bfed892dbd4c4d9f9871b3b1f05325ff"
        Me.OpenItemFromNote.ImageMso = "ShowNotesPage"
        Me.OpenItemFromNote.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.OpenItemFromNote.Ribbons = CType((AddinExpress.MSO.ADXRibbons.msrOutlookAppointment Or AddinExpress.MSO.ADXRibbons.msrOutlookTask), AddinExpress.MSO.ADXRibbons)
        Me.OpenItemFromNote.ScreenTip = "Open a NewCall Tracking or Appointment item from the Note that is attached to the" & _
    " currently displayed item."
        Me.OpenItemFromNote.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonSeparator2
        '
        Me.AdxRibbonSeparator2.Id = "adxRibbonSeparator_3b3486f982d94626afeeaacb820d8686"
        Me.AdxRibbonSeparator2.Ribbons = CType((AddinExpress.MSO.ADXRibbons.msrOutlookAppointment Or AddinExpress.MSO.ADXRibbons.msrOutlookTask), AddinExpress.MSO.ADXRibbons)
        '
        'AdxRibbonButton1
        '
        Me.AdxRibbonButton1.Caption = "Copy Item to Drafts Folder"
        Me.AdxRibbonButton1.Id = "adxRibbonButton_de572db5435f4b35bc780bc9d332c327"
        Me.AdxRibbonButton1.ImageMso = "SendStatusReport"
        Me.AdxRibbonButton1.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AdxRibbonButton1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookTask
        Me.AdxRibbonButton1.ScreenTip = "Copy this item to your Drafts folder (so you can E-mail a copy of it to someone)." & _
    ""
        Me.AdxRibbonButton1.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'CopyAttachments
        '
        Me.CopyAttachments.Caption = "Copy Attachments"
        Me.CopyAttachments.Id = "adxRibbonButton_e906c4b7b7734aa0bcf88b14b7ded48c"
        Me.CopyAttachments.ImageMso = "AttachItem"
        Me.CopyAttachments.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.CopyAttachments.Ribbons = CType((((AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose Or AddinExpress.MSO.ADXRibbons.msrOutlookResend) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookPostCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSCompose), AddinExpress.MSO.ADXRibbons)
        Me.CopyAttachments.ScreenTip = "Copy the attachments from another E-mail to this E-mail."
        Me.CopyAttachments.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonSeparator3
        '
        Me.AdxRibbonSeparator3.Id = "adxRibbonSeparator_49791e835dbc46a2be8c57236b4b548c"
        Me.AdxRibbonSeparator3.Ribbons = CType(((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
        '
        'AdxRibbonButton4
        '
        Me.AdxRibbonButton4.Caption = "About"
        Me.AdxRibbonButton4.Id = "adxRibbonButton_54403c6dbcd54328a298f2556650e611"
        Me.AdxRibbonButton4.ImageMso = "Help"
        Me.AdxRibbonButton4.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AdxRibbonButton4.Ribbons = CType((((((((((((((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMeetingRequestRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMeetingRequestSend) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookResend) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookResponseRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookResponseCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookResponseCounterPropose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookPostRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookPostCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookSharingRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookSharingCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSCompose), AddinExpress.MSO.ADXRibbons)
        Me.AdxRibbonButton4.ScreenTip = "Display information about this ribbon tab."
        Me.AdxRibbonButton4.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'OpenApptFromFile
        '
        Me.OpenApptFromFile.Caption = "Open Note from File"
        Me.OpenApptFromFile.Id = "adxRibbonButton_fd25b74e412c4e36a85c28c29d6c6e1a"
        Me.OpenApptFromFile.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.OpenApptFromFile.Ribbons = AddinExpress.MSO.ADXRibbons.msrNone
        Me.OpenApptFromFile.ScreenTip = "Open the test Appointment."
        Me.OpenApptFromFile.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        Me.OpenApptFromFile.SuperTip = "Opens the test Appointment from ""C:\tmp\NewCall Appointment.msg"""
        Me.OpenApptFromFile.Visible = False
        '
        'AdxOutlookAppEvents1
        '
        '
        'AdxRibbonSeparator4
        '
        Me.AdxRibbonSeparator4.Id = "adxRibbonSeparator_81545993843d4e66aa25cab258f35721"
        Me.AdxRibbonSeparator4.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookTask
        '
        'AdxRibbonSeparator5
        '
        Me.AdxRibbonSeparator5.Id = "adxRibbonSeparator_d30c2775514e4ff7960047d7e3ba2371"
        Me.AdxRibbonSeparator5.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookContact
        '
        'AddinModule
        '
        Me.AddinName = "GKBMOutlook"
        Me.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook

    End Sub
    Private WithEvents AdxRibbonTab1 As AddinExpress.MSO.ADXRibbonTab
    Friend WithEvents AdxRibbonGroup1 As AddinExpress.MSO.ADXRibbonGroup
    Friend WithEvents CopyContact2InstantFile As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents AdxRibbonButton2 As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents AdxRibbonButtonSaveAttachments As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents AdxRibbonButton1 As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents AdxRibbonButton4 As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents SaveClose As AddinExpress.MSO.ADXRibbonButton
    Private WithEvents AdxOutlookAppEvents1 As AddinExpress.MSO.ADXOutlookAppEvents
    Friend WithEvents CopyAttachments As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents OpenApptFromFile As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents OpenItemFromNote As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents MakeAppointment As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents AdxRibbonSeparator1 As AddinExpress.MSO.ADXRibbonSeparator
    Friend WithEvents AdxRibbonSeparator2 As AddinExpress.MSO.ADXRibbonSeparator
    Friend WithEvents AdxRibbonSeparator3 As AddinExpress.MSO.ADXRibbonSeparator
    Friend WithEvents AdxRibbonSeparator5 As AddinExpress.MSO.ADXRibbonSeparator
    Friend WithEvents AdxRibbonSeparator4 As AddinExpress.MSO.ADXRibbonSeparator

End Class

