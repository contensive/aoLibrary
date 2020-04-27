
Namespace Models.Domain
    '
    Public Class FolderTypeModel
        Public FolderID As Integer
        Public parentFolderID As Integer
        Public Name As String
        Public FullPath As String
        '
        Public hasViewAccess As Boolean                    ' has permission to view this folder (below topFolderPath)
        Public viewAccessIsValid As Boolean                 ' true when hasViewAccess is correct
        '
        Public hasModifyAccess As Boolean                  ' has permission to modify files and folders in this folder
        Public modifyAccessIsValid As Boolean              ' true when hasModifyAccess is correct
    End Class
End Namespace
