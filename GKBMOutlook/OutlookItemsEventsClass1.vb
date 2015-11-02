Imports System
Imports System.Windows.Forms
Imports AddinExpress.MSO

'Add-in Express Outlook Items Events Class
Public Class OutlookItemsEventsClass1
    Inherits AddinExpress.MSO.ADXOutlookItemsEvents
 
    Public Sub New(ByVal ADXModule As AddinExpress.MSO.ADXAddinModule)
        MyBase.New(ADXModule)
 
    End Sub
 
    Public Overrides Sub ItemAdd(ByVal Item As Object, ByVal SourceFolder As Object)
        MsgBox("ItemAdd fired.")
    End Sub
 
    Public Overrides Sub ItemChange(ByVal Item As Object, ByVal SourceFolder As Object)
        'TODO: Add some code
    End Sub
 
    Public Overrides Sub ItemRemove(ByVal SourceFolder As Object)
        'TODO: Add some code
    End Sub
 
    Public Overrides Sub BeforeFolderMove(ByVal moveTo As Object, ByVal SourceFolder As Object, ByVal e As AddinExpress.MSO.ADXCancelEventArgs)
        'TODO: Add some code
    End Sub
 
    Public Overrides Sub BeforeItemMove(ByVal item As Object, ByVal moveTo As Object, ByVal SourceFolder As Object, ByVal e As AddinExpress.MSO.ADXCancelEventArgs)
        'TODO: Add some code
    End Sub
 
End Class

