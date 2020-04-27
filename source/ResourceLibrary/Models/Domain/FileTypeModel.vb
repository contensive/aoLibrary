
Namespace Models.Domain
    '
    Public Class FileTypeModel
        Public Name As String = ""
        Public FileTypeID As Integer
        Public ExtensionList As String = ""
        Public IconFilename As String = ""
        Public IsImage As Boolean
        Public IsFlash As Boolean
        Public IsVideo As Boolean
        Public MediaIconFilename As String = ""
        Public IsDownload As Boolean
        Public DownloadIconFilename As String = ""
    End Class
End Namespace
