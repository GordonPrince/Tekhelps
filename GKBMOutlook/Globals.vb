Imports System.Data  ' includes SqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Module Globals
    Public Const strDocScanned As String = "Document scanned + imported:"
    Public Const strLastScanned As String = "LAST REQUESTED DOCUMENT scanned + imported:"
    Public Const strIFmatNo As String = "InstantFile_MatNo_"
    Public Const strIFdocNo As String = "InstantFile_DocNo_"
    Public Const strPublicFolders As String = "Public Folders"
    Public Const strAllPublicFolders As String = "All Public Folders"
    Public Const strInstantFile As String = "InstantFile"
    Public Const strNewCallTrackingTag As String = "NewCall Tracking Item"
    Public Const strNewCallAppointmentTag As String = "NewCall Appointment"
    Public Const strIFtaskTag As String = "InstantFile_Task"
    Public Const strSend2Gordon As String = vbNewLine & vbNewLine & "Please use the Snipping Tool to capture this message" & vbNewLine & _
                                                "and E-mail it to Gordon."

    Public strScratch As String
    Public strPublicStoreID As String
    Public RetVal As VariantType
    Public lngX As Long

    Public myPublicFolder As Outlook.Folder = Nothing

    ' Private OutlookApp As Outlook.Application = CType(AddinModule.CurrentInstance, AddinModule).OutlookApp
    Public OutlookApp As Outlook.Application = Nothing

    Public Function RunSQLcommand(ByVal queryString As String) As Boolean
        Dim strConnectionString As String = SQLConnectionString()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        Dim cmd As New SqlClient.SqlCommand(queryString, con)
        ' Using con As New SqlClient.SqlConnection(strConnectionString)
        Try
            cmd.Connection.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
        ' End Using
    End Function

    Public Function SQLConnectionString() As String
        If My.Computer.Name = "TEKHELPS7X64" Then
            SQLConnectionString = ("Initial Catalog=InstantFile;Data Source=TEKHELPS7X64\SQL2005X64;Integrated Security=SSPI;")
        Else
            SQLConnectionString = ("Initial Catalog=InstantFile;Data Source=SQLserver;Integrated Security=SSPI;")
        End If
    End Function

    Public Function GetPublicFolder(ByVal strFolderName As String) As Boolean ' 11/17/2015
        Dim mySession As Outlook.NameSpace = Nothing
        Dim myFolders As Outlook.Folders = Nothing
        Dim myFolder As Outlook.Folder = Nothing

        ' also try myFolder = myFolders.GetNext 
        ' https://msdn.microsoft.com/en-us/library/office/ff865587.aspx?f=255&MSPPError=-2147217396
        Try
            mySession = OutlookApp.Session
            myFolders = mySession.Folders
            If myFolders.Count > 0 Then
                Dim x As Short
                For x = 1 To myFolders.Count
                    myFolder = myFolders(x)
                    If Left(myFolder.Name, Len(strPublicFolders)) = strPublicFolders Then
                        Marshal.ReleaseComObject(myFolders)
                        myFolders = myFolder.Folders
                        Marshal.ReleaseComObject(myFolder)
                        Dim y As Short
                        For y = 1 To myFolders.Count
                            myFolder = myFolders(y)
                            If myFolder.Name = strAllPublicFolders Then
                                'Marshal.ReleaseComObject(myFolders)
                                myFolders = myFolder.Folders
                                Marshal.ReleaseComObject(myFolder)
                                myFolder = myFolders.GetFirst
                                Do While Not myFolder Is Nothing
                                    If myFolder.Name = strFolderName Then
                                        If myPublicFolder IsNot Nothing Then Marshal.ReleaseComObject(myPublicFolder)
                                        myPublicFolder = myFolder
                                        Return True
                                    Else
                                        myFolder = myFolders.GetNext
                                    End If
                                Loop
                                Return False
                            End If
                        Next ' Public Folders
                    End If
                Next ' Session folders
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "GetPublicFolder()")
        Finally
            ' don't release myFolder -- it's the same object as myPublicFolder
            ' If myFolder IsNot Nothing Then Marshal.ReleaseComObject(myFolder) : myFolder = Nothing
            If myFolders IsNot Nothing Then Marshal.ReleaseComObject(myFolders) : myFolders = Nothing
            If mySession IsNot Nothing Then Marshal.ReleaseComObject(mySession) : mySession = Nothing
        End Try
        Return False
    End Function
End Module
