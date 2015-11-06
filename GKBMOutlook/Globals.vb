Imports System.Data

Module Globals
    Public Const strDocScanned As String = "Document scanned + imported:"
    Public Const strLastScanned As String = "LAST REQUESTED DOCUMENT scanned + imported:"
    Public Const strIFmatNo As String = "InstantFile_MatNo_"
    Public Const strIFdocNo As String = "InstantFile_DocNo_"
    Public Const strPublicFolders As String = "Public Folders"
    Public Const strAllPublicFolders As String = "All Public Folders"
    Public Const strInstantFile As String = "InstantFile"
    Public Const strNewCallTrackingTag As String = "NewCall Tracking Item"
    Public Const strIFtaskTag As String = "InstantFile_Task"
    Public Const strNewCallAppointmentTag As String = "NewCall Appointment"

    Public strScratch As String
    Public strPublicStoreID As String
    Public RetVal As VariantType
    Public lngX As Long

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

End Module
