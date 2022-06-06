Imports System.Runtime.InteropServices

Public Enum KnowFolder
    Contacts
    Downloads
    Favorites
    Links
    SavedSearches
End Enum


Public Class KnowFolders
    Private Shared ReadOnly FolderGuids As New Dictionary(Of KnowFolder, Guid) From {
        {KnowFolder.Contacts, New Guid("56784854-C6CB-462B-8169-88E350ACB882")},
        {KnowFolder.Downloads, New Guid("374DE290-123F-4565-9164-39C4925E467B")},
        {KnowFolder.Favorites, New Guid("1777F761-68AD-4D8A-87BD-30B759FA33DD")},
        {KnowFolder.Links, New Guid("BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968")},
        {KnowFolder.SavedSearches, New Guid("7D1D3A04-DEBB-4115-95CF-2F29DA2920DA")}
    }

    <DllImport("shell32.dll")>
    Private Shared Function SHGetKnownFolderPath(<MarshalAs(UnmanagedType.LPStruct)> ByVal rfid As Guid,
                                                  ByVal dwFlags As UInteger,
                                                  ByVal hToken As IntPtr,
                                                  ByRef pszPath As IntPtr) As Integer
    End Function

    Public Shared Function GetKnowFolder(guid As KnowFolder) As String
        Dim Result As String = ""
        Dim ppszPath As IntPtr
        Dim gGuid As Guid = FolderGuids(guid)

        If SHGetKnownFolderPath(gGuid, 0, 0, ppszPath) = 0 Then
            Result = Marshal.PtrToStringUni(ppszPath)
            Marshal.FreeCoTaskMem(ppszPath)
        End If
        Return Result
    End Function
End Class
