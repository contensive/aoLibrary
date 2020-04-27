

Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Models.Db
    Public Class ResourceLibraryModel
        Inherits DesignBlockBaseModel
        '
        '====================================================================================================
        ''' <summary>
        ''' table definition
        ''' </summary>
        Public Shared ReadOnly Property tableMetadata As DbBaseTableMetadataModel = New DbBaseTableMetadataModel("Resource Libraries", "ccResourceLibraries", "default", False)
        '
        '====================================================================================================
        '
        Public Property RootFolderName As String
        'Public Property AllowGroupAdd As Boolean
        'Public Property AllowSelectResource As Boolean
        'Public Property SelectResourceEditorObjectName As Boolean
        'Public Property SelectLinkObjectName As String
        Public Property BlockFolderNavigation As Boolean
        '
        '====================================================================================================
        Public Overloads Shared Function createOrAddSettings(cp As CPBaseClass, settingsGuid As String) As ResourceLibraryModel
            Dim result As ResourceLibraryModel = create(Of ResourceLibraryModel)(cp, settingsGuid)
            If (result Is Nothing) Then
                '
                ' -- create default content
                result = DesignBlockBaseModel.addDefault(Of ResourceLibraryModel)(cp)
                result.name = ResourceLibraryModel.tableMetadata.contentName & " " & result.id
                result.ccguid = settingsGuid
                'result.fontStyleId = 0
                result.themeStyleId = 0
                result.padTop = False
                result.padBottom = False
                result.padRight = False
                result.padLeft = False
                '
                ' -- create custom content
                result.RootFolderName = ""
                result.BlockFolderNavigation = False
                '
                result.save(cp)
                '
                ' -- track the last modified date
                cp.Content.LatestContentModifiedDate.Track(result.modifiedDate)
                '
            End If
            Return result
        End Function
    End Class
End Namespace