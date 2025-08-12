Attribute VB_Name = "funcOneDrive"

Public Function OneDrivePath(S_Url As String) As String

    Const Cns_OneDriveCommercialPattern As String = "my.sharepoint.com"
    ' Right-hand value for Like operator to determine if the URL is for OneDrive for Business

    Dim S_pathSeparator As String
    Dim S_OneDriveCommercialPath As String
    Dim S_OneDriveConsumerPath As String

    Dim S_PathPosition As Long

    'If the argument is not a URL, assume it is a local path and return it as is.
    If Not (S_Url Like "https://*") Then
        OneDrivePath = S_Url
        Exit Function
    End If

    S_pathSeparator = Application.PathSeparator

    'Path for OneDrive for Business (Commercial)
    S_OneDriveCommercialPath = Environ("OneDriveCommercial")
    If (S_OneDriveCommercialPath = "") Then
        S_OneDriveCommercialPath = Environ("OneDrive")
    End If

    ' Path for personal OneDrive
    S_OneDriveConsumerPath = Environ("OneDriveConsumer")
    If (S_OneDriveConsumerPath = "") Then
        S_OneDriveConsumerPath = Environ("OneDrive")
    End If

    ' For business OneDrive: S_Url = "https://[company]-my.sharepoint.com/personal/[username]_domain_com/Documents/[filepath]"
    If (S_Url Like "*" & Cns_OneDriveCommercialPattern & "*") Then

        S_PathPosition = InStr(1, S_Url, "/Documents") + 10                         '10 = Len("/Documents")
        OneDrivePath = S_OneDriveCommercialPath & Replace(Mid(S_Url, S_PathPosition), "/", S_pathSeparator)

    ' For personal OneDrive: S_Url = "https://d.docs.live.net/[CID_number]/[file_path]"
    Else
        '********************************************************************************
        '         1         2         3         4         5         6         7         8
        '12345678901234567890123456789012345678901234567890123456789012345678901234567890
        'https://d.docs.live.net/f53c0b88b096e170/desktop/Excel_Ichiran_Ver3.4
        '********************************************************************************

        S_PathPosition = InStr(9, S_Url, "/")                                       '9 = Len("https://") + 1
        S_PathPosition = InStr(S_PathPosition + 1, S_Url, "/")
        OneDrivePath = S_OneDriveConsumerPath & Replace(Mid(S_Url, S_PathPosition), "/", S_pathSeparator)

    End If

End Function

Public Sub tmp()

'This command returns the URL for files on OneDrive.
    Debug.Print ThisWorkbook.Path
    Debug.Print ActiveWorkbook.Path

'This command returns the file path in any environment.
    Debug.Print OneDrivePath(ThisWorkbook.Path)
    Debug.Print OneDrivePath(ActiveWorkbook.Path)

End Sub


