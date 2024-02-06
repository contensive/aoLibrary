
Imports Contensive.Addons.ResourceLibrary.Controllers
Imports Contensive.Addons.ResourceLibrary.Models.View
Imports Contensive.BaseClasses

Namespace Views
    '
    '====================================================================================================
    ''' <summary>
    ''' Design block with a centered headline, image, paragraph text and a button.
    ''' </summary>
    Public Class LibraryClass
        Inherits AddonBaseClass
        '
        '====================================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Const designBlockName As String = "Resource Library"
            Try
                Dim settings As Models.Db.ResourceLibraryModel
                If (CP.Doc.IsAdminSite) Then
                    '
                    ' -- admin site settings
                    settings = New Models.Db.ResourceLibraryModel() With {
                        .RootFolderName = "",
                        .BlockFolderNavigation = False
                    }
                Else
                    '
                    ' -- design block settings
                    Dim userErrorMessage As String = ""
                    Dim settingsGuid = InstanceIdController.getSettingsGuid(CP, designBlockName, userErrorMessage)
                    If (String.IsNullOrEmpty(settingsGuid)) Then Return userErrorMessage
                    '
                    ' -- locate or create a data record for this guid
                    settings = Models.Db.ResourceLibraryModel.createOrAddSettings(CP, settingsGuid)
                    If (settings Is Nothing) Then Throw New ApplicationException("Could not create the design block settings record.")
                    '
                    CP.Doc.SetProperty("RootFolderName", settings.RootFolderName)
                    CP.Doc.SetProperty("Block Folder Navigation", settings.BlockFolderNavigation)

                End If
                Dim htmlBody = (New LegacyLibraryClass).getResourceLibrary(CP)
                '
                ' -- translate the Db model to a view model and mustache it into the layout
                Dim viewModel = LibraryViewModel.create(CP, settings, htmlBody)
                If (viewModel Is Nothing) Then Throw New ApplicationException("Could not create design block view model.")
                Dim result As String = Nustache.Core.Render.StringToString(My.Resources.LibraryLayout, viewModel)
                '
                ' -- if editing enabled, add the link and wrapperwrapper
                Return CP.Content.GetEditWrapper(result, settings.name, settings.id)
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
                Return "<!-- " & designBlockName & ", Unexpected Exception -->"
            End Try
        End Function
    End Class
End Namespace