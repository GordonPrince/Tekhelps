Imports System.Windows.Forms

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
        Me.CopyContact2InstantFile = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonButton2 = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonButtonSaveAttachments = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonButton1 = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonButton4 = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        '
        'AdxRibbonTab1
        '
        Me.AdxRibbonTab1.Caption = "GKBM"
        Me.AdxRibbonTab1.Controls.Add(Me.AdxRibbonGroup1)
        Me.AdxRibbonTab1.Id = "adxRibbonTab_a5cf84a1ced7485ea57b9386db6cf9ed"
        Me.AdxRibbonTab1.Ribbons = CType(((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
        '
        'AdxRibbonGroup1
        '
        Me.AdxRibbonGroup1.Caption = "Custom Functions"
        Me.AdxRibbonGroup1.Controls.Add(Me.SaveClose)
        Me.AdxRibbonGroup1.Controls.Add(Me.CopyContact2InstantFile)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButton2)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButtonSaveAttachments)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButton1)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButton4)
        Me.AdxRibbonGroup1.Id = "adxRibbonGroup_072621b7e27f4dc6966496c4216ab446"
        Me.AdxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AdxRibbonGroup1.Ribbons = CType(((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
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
        'CopyContact2InstantFile
        '
        Me.CopyContact2InstantFile.Caption = "Add Contact to InstantFile"
        Me.CopyContact2InstantFile.Id = "adxRibbonButton_0e8a15c449ee4798adf475a1587fd047"
        Me.CopyContact2InstantFile.ImageMso = "ContactCardInstantMessageAddToOutlookContacts"
        Me.CopyContact2InstantFile.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.CopyContact2InstantFile.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookContact
        Me.CopyContact2InstantFile.ScreenTip = "Copy this personal Contact to InstantFile's Contacts"
        Me.CopyContact2InstantFile.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonButton2
        '
        Me.AdxRibbonButton2.Caption = "Link Two Contacts"
        Me.AdxRibbonButton2.Id = "adxRibbonButton_4b812ad786fc4beb9f6a39f3bbcb0352"
        Me.AdxRibbonButton2.ImageMso = "ContactRoles"
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
        'AdxRibbonButton1
        '
        Me.AdxRibbonButton1.Caption = "Copy Item to Drafts Folder"
        Me.AdxRibbonButton1.Id = "adxRibbonButton_de572db5435f4b35bc780bc9d332c327"
        Me.AdxRibbonButton1.ImageMso = "ForwardTask"
        Me.AdxRibbonButton1.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AdxRibbonButton1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookTask
        Me.AdxRibbonButton1.ScreenTip = "Copy this item to your Drafts folder (so you can E-mail a copy of it to someone)." & _
    ""
        Me.AdxRibbonButton1.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonButton4
        '
        Me.AdxRibbonButton4.Caption = "About"
        Me.AdxRibbonButton4.Id = "adxRibbonButton_54403c6dbcd54328a298f2556650e611"
        Me.AdxRibbonButton4.ImageMso = "Help"
        Me.AdxRibbonButton4.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AdxRibbonButton4.Ribbons = CType((((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMeetingRequestRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookPostRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSRead), AddinExpress.MSO.ADXRibbons)
        Me.AdxRibbonButton4.ScreenTip = "Display information about this ribbon tab."
        Me.AdxRibbonButton4.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
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

End Class

