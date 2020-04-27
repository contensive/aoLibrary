
Imports Contensive.BaseClasses

Namespace Models.Db
    Public Class LibraryFolderModel
        Inherits Global.Contensive.Models.Db.LibraryFolderModel


        '''' <summary>
        '''' Return a list of folders
        '''' </summary>
        '''' <param name="cp"></param>
        '''' <param name="FolderID">The id of the folder</param>
        '''' <returns></returns>
        'Public Shared Function AllowFolderAccess(cp As CPBaseClass, FolderID As Integer, ParentID As Integer) As List(Of LibraryFolderModel)
        '    Dim result As New List(Of LibraryFolderModel)
        '    Try
        '        Dim SQL = "select top 1 *" _
        '              & " from ccMemberRules M,ccLibraryFolderRules R" _
        '              & " where M.MemberID=" & cp.User.Id _
        '              & " and R.FolderID=" & FolderID _
        '              & " and M.GroupID=R.GroupID" _
        '              & " and R.Active<>0" _
        '              & " and M.Active<>0" _
        '              & " and ((M.DateExpires is null)or(M.DateExpires>" & cp.Db.EncodeSQLDate(Now) & "))"
        '        result = createList(Of LibraryFolderModel)(cp, "(FolderID in (" & SQL & "))", "name")
        '    Catch ex As Exception
        '        cp.Site.ErrorReport(ex)
        '    End Try
        '    Return result
        'End Function
    End Class
End Namespace