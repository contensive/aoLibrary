
Imports Contensive.BaseClasses


Namespace Models.View
    Public Class LibraryViewModel
        Inherits DesignBlockViewBaseModel
        '
        Public Property bodyHtml As String
        '
        '====================================================================================================
        ''' <summary>
        ''' Populate the view model from the entity model
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="settings"></param>
        ''' <returns></returns>
        Public Overloads Shared Function create(cp As CPBaseClass, settings As Models.Db.ResourceLibraryModel, htmlBody As String) As LibraryViewModel
            Try
                '
                ' -- base fields
                Dim result = DesignBlockViewBaseModel.create(Of LibraryViewModel)(cp, settings)
                '
                ' -- custom
                result.bodyHtml = htmlBody
                '
                Return result
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Return Nothing
            End Try
        End Function
    End Class

End Namespace