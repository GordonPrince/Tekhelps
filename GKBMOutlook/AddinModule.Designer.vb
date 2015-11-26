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
        Me.AdxRibbonGroup2 = New AddinExpress.MSO.ADXRibbonGroup(Me.components)
        Me.SaveClose = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonSeparator6 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        Me.AppointmentCalendar = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonSeparator2 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        Me.SSIcalendar = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonSeparator3 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        Me.NewCallTracking = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonGroup1 = New AddinExpress.MSO.ADXRibbonGroup(Me.components)
        Me.CopyContact2InstantFile = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonSeparator5 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        Me.AdxRibbonButton2 = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonButtonSaveAttachments = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.MakeAppointment = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonSeparator4 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        Me.OpenItemFromNote = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxRibbonSeparator1 = New AddinExpress.MSO.ADXRibbonSeparator(Me.components)
        Me.AdxRibbonButton1 = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.CopyAttachments = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.About = New AddinExpress.MSO.ADXRibbonGroup(Me.components)
        Me.AdxRibbonButton4 = New AddinExpress.MSO.ADXRibbonButton(Me.components)
        Me.AdxOutlookAppEvents1 = New AddinExpress.MSO.ADXOutlookAppEvents(Me.components)
        '
        'AdxRibbonTab1
        '
        Me.AdxRibbonTab1.Caption = "GKBM"
        Me.AdxRibbonTab1.Controls.Add(Me.AdxRibbonGroup2)
        Me.AdxRibbonTab1.Controls.Add(Me.AdxRibbonGroup1)
        Me.AdxRibbonTab1.Controls.Add(Me.About)
        Me.AdxRibbonTab1.Id = "adxRibbonTab_a5cf84a1ced7485ea57b9386db6cf9ed"
        Me.AdxRibbonTab1.InsertBeforeIdMso = "TabInsert"
        Me.AdxRibbonTab1.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
        '
        'AdxRibbonGroup2
        '
        Me.AdxRibbonGroup2.Caption = "Items"
        Me.AdxRibbonGroup2.Controls.Add(Me.SaveClose)
        Me.AdxRibbonGroup2.Controls.Add(Me.AdxRibbonSeparator6)
        Me.AdxRibbonGroup2.Controls.Add(Me.AppointmentCalendar)
        Me.AdxRibbonGroup2.Controls.Add(Me.AdxRibbonSeparator2)
        Me.AdxRibbonGroup2.Controls.Add(Me.SSIcalendar)
        Me.AdxRibbonGroup2.Controls.Add(Me.AdxRibbonSeparator3)
        Me.AdxRibbonGroup2.Controls.Add(Me.NewCallTracking)
        Me.AdxRibbonGroup2.Id = "adxRibbonGroup_9810b0e61d06422397f5d3ddba5337a1"
        Me.AdxRibbonGroup2.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AdxRibbonGroup2.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
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
        Me.SaveClose.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.SaveClose.Ribbons = CType(((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose Or AddinExpress.MSO.ADXRibbons.msrOutlookMeetingRequestRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookPostRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSRead), AddinExpress.MSO.ADXRibbons)
        Me.SaveClose.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonSeparator6
        '
        Me.AdxRibbonSeparator6.Id = "adxRibbonSeparator_b6799511e297445ba5bac1a95c3d48bd"
        Me.AdxRibbonSeparator6.Ribbons = CType(((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose Or AddinExpress.MSO.ADXRibbons.msrOutlookMeetingRequestRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookPostRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSRead), AddinExpress.MSO.ADXRibbons)
        '
        'AppointmentCalendar
        '
        Me.AppointmentCalendar.Caption = "Appointment Calendar"
        Me.AppointmentCalendar.Id = "adxRibbonButton_0bba54f1190c4c62b8f84191be4195a2"
        Me.AppointmentCalendar.ImageMso = "OpenAttachedCalendar"
        Me.AppointmentCalendar.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AppointmentCalendar.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
        Me.AppointmentCalendar.ScreenTip = "Display the Appointment Calendar."
        Me.AppointmentCalendar.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonSeparator2
        '
        Me.AdxRibbonSeparator2.Id = "adxRibbonSeparator_9f533fe0de0d4a9e9f968cf4bfbb8ee6"
        Me.AdxRibbonSeparator2.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
        '
        'SSIcalendar
        '
        Me.SSIcalendar.Caption = "SSI Calendar"
        Me.SSIcalendar.Id = "adxRibbonButton_d19abb57f9e14dd09b8abf7391a9134d"
        Me.SSIcalendar.ImageMso = "MeetingsWorkspace"
        Me.SSIcalendar.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.SSIcalendar.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
        Me.SSIcalendar.ScreenTip = "Display the Social Security Appointment Calendar."
        Me.SSIcalendar.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonSeparator3
        '
        Me.AdxRibbonSeparator3.Id = "adxRibbonSeparator_c517d1560742417d97db622349f06e0b"
        Me.AdxRibbonSeparator3.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
        '
        'NewCallTracking
        '
        Me.NewCallTracking.Caption = "New Call Tracking"
        Me.NewCallTracking.Id = "adxRibbonButton_e078a19a82914dadbdcd8f0242771ea6"
        Me.NewCallTracking.ImageMso = "AccessListContacts"
        Me.NewCallTracking.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.NewCallTracking.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
        Me.NewCallTracking.ScreenTip = "Display the New Call Tracking folder."
        Me.NewCallTracking.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonGroup1
        '
        Me.AdxRibbonGroup1.Caption = "Functions"
        Me.AdxRibbonGroup1.Controls.Add(Me.CopyContact2InstantFile)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonSeparator5)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButton2)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButtonSaveAttachments)
        Me.AdxRibbonGroup1.Controls.Add(Me.MakeAppointment)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonSeparator4)
        Me.AdxRibbonGroup1.Controls.Add(Me.OpenItemFromNote)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonSeparator1)
        Me.AdxRibbonGroup1.Controls.Add(Me.AdxRibbonButton1)
        Me.AdxRibbonGroup1.Controls.Add(Me.CopyAttachments)
        Me.AdxRibbonGroup1.Id = "adxRibbonGroup_072621b7e27f4dc6966496c4216ab446"
        Me.AdxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.AdxRibbonGroup1.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookContact) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer), AddinExpress.MSO.ADXRibbons)
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
        'AdxRibbonSeparator5
        '
        Me.AdxRibbonSeparator5.Id = "adxRibbonSeparator_d30c2775514e4ff7960047d7e3ba2371"
        Me.AdxRibbonSeparator5.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookContact
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
        Me.MakeAppointment.Caption = "Make New Appointment"
        Me.MakeAppointment.Id = "adxRibbonButton_6125721066004d2a94e0fb7b2a58af6f"
        Me.MakeAppointment.ImageMso = "NewMeetingRequestNumbered"
        Me.MakeAppointment.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.MakeAppointment.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookTask
        Me.MakeAppointment.ScreenTip = "Make a New Appointment for this Caller."
        Me.MakeAppointment.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonSeparator4
        '
        Me.AdxRibbonSeparator4.Id = "adxRibbonSeparator_81545993843d4e66aa25cab258f35721"
        Me.AdxRibbonSeparator4.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookTask
        '
        'OpenItemFromNote
        '
        Me.OpenItemFromNote.Caption = "Open Item from Note"
        Me.OpenItemFromNote.Id = "adxRibbonButton_bfed892dbd4c4d9f9871b3b1f05325ff"
        Me.OpenItemFromNote.ImageMso = "ShowNotesPage"
        Me.OpenItemFromNote.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.OpenItemFromNote.Ribbons = CType(((((((((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMeetingRequestRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMeetingRequestSend) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookTask) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookResend) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookResponseRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookResponseCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookResponseCounterPropose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookExplorer) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSRead) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookMMSCompose), AddinExpress.MSO.ADXRibbons)
        Me.OpenItemFromNote.ScreenTip = "Open a NewCall Tracking or Appointment item from the Note that is attached to the" & _
    " currently displayed item."
        Me.OpenItemFromNote.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'AdxRibbonSeparator1
        '
        Me.AdxRibbonSeparator1.Id = "adxRibbonSeparator_ab15ec8bc2634c16b73d15434c77cde1"
        Me.AdxRibbonSeparator1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookTask
        '
        'AdxRibbonButton1
        '
        Me.AdxRibbonButton1.Caption = "E-mail Copy of This Item"
        Me.AdxRibbonButton1.Id = "adxRibbonButton_de572db5435f4b35bc780bc9d332c327"
        Me.AdxRibbonButton1.ImageMso = "FileSendAsAttachment"
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
        Me.CopyAttachments.ScreenTip = "Copy the attachments from other open E-mail(s) to this E-mail."
        Me.CopyAttachments.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large
        '
        'About
        '
        Me.About.Caption = "Help"
        Me.About.Controls.Add(Me.AdxRibbonButton4)
        Me.About.Id = "adxRibbonGroup_5a2f9b298cd04e71bac140e4dac5c235"
        Me.About.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.About.Ribbons = CType((((((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead Or AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose) _
            Or AddinExpress.MSO.ADXRibbons.msrOutlookAppointment) _
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
        'AdxOutlookAppEvents1
        '
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
    Friend WithEvents OpenItemFromNote As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents MakeAppointment As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents AdxRibbonSeparator5 As AddinExpress.MSO.ADXRibbonSeparator
    Friend WithEvents AdxRibbonSeparator4 As AddinExpress.MSO.ADXRibbonSeparator
    Friend WithEvents NewCallTracking As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents AppointmentCalendar As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents AdxRibbonGroup2 As AddinExpress.MSO.ADXRibbonGroup
    Friend WithEvents About As AddinExpress.MSO.ADXRibbonGroup
    Friend WithEvents AdxRibbonSeparator1 As AddinExpress.MSO.ADXRibbonSeparator
    Friend WithEvents SSIcalendar As AddinExpress.MSO.ADXRibbonButton
    Friend WithEvents AdxRibbonSeparator6 As AddinExpress.MSO.ADXRibbonSeparator
    Friend WithEvents AdxRibbonSeparator2 As AddinExpress.MSO.ADXRibbonSeparator
    Friend WithEvents AdxRibbonSeparator3 As AddinExpress.MSO.ADXRibbonSeparator

End Class

