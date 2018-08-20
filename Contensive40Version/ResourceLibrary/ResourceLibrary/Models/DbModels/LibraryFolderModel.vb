
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Contensive.Addons.ResourceLibrary.Models     '<------ set namespace
    Public Class LibraryFolderModel        '<------ set set model Name and everywhere that matches this string
        Inherits baseModel
        Implements ICloneable
        '
        '====================================================================================================
        '-- const
        Public Const contentName As String = "Library Folders"
        Public Const contentTableName As String = "ccLibraryFolders"
        Private Shadows Const contentDataSource As String = "default"
        '
        '====================================================================================================
        ' -- instance properties
        Public Property Description As String
        Public Property ParentID As Integer
        '
        '====================================================================================================
        Public Overloads Shared Function add(cp As CPBaseClass) As LibraryFolderModel
            Return add(Of LibraryFolderModel)(cp)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function create(cp As CPBaseClass, recordId As Integer) As LibraryFolderModel
            Return create(Of LibraryFolderModel)(cp, recordId)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function create(cp As CPBaseClass, recordGuid As String) As LibraryFolderModel
            Return create(Of LibraryFolderModel)(cp, recordGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function createByName(cp As CPBaseClass, recordName As String) As LibraryFolderModel
            Return createByName(Of LibraryFolderModel)(cp, recordName)
        End Function
        '
        '====================================================================================================
        Public Overloads Sub save(cp As CPBaseClass)
            MyBase.save(Of LibraryFolderModel)(cp)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Sub delete(cp As CPBaseClass, recordId As Integer)
            delete(Of LibraryFolderModel)(cp, recordId)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Sub delete(cp As CPBaseClass, ccGuid As String)
            delete(Of LibraryFolderModel)(cp, ccGuid)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Function createList(cp As CPBaseClass, sqlCriteria As String, Optional sqlOrderBy As String = "id") As List(Of LibraryFolderModel)
            Return createList(Of LibraryFolderModel)(cp, sqlCriteria, sqlOrderBy)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordName(cp As CPBaseClass, recordId As Integer) As String
            Return baseModel.getRecordName(Of LibraryFolderModel)(cp, recordId)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordName(cp As CPBaseClass, ccGuid As String) As String
            Return baseModel.getRecordName(Of LibraryFolderModel)(cp, ccGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordId(cp As CPBaseClass, ccGuid As String) As Integer
            Return baseModel.getRecordId(Of LibraryFolderModel)(cp, ccGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getCount(cp As CPBaseClass, sqlCriteria As String) As Integer
            Return baseModel.getCount(Of LibraryFolderModel)(cp, sqlCriteria)
        End Function
        '
        '====================================================================================================
        Public Overloads Function getUploadPath(fieldName As String) As String
            Return MyBase.getUploadPath(Of LibraryFolderModel)(fieldName)
        End Function
        '
        '====================================================================================================
        '
        Public Function Clone(cp As CPBaseClass) As LibraryFolderModel
            Dim result As LibraryFolderModel = DirectCast(Me.Clone(), LibraryFolderModel)
            result.id = cp.Content.AddRecord(contentName)
            result.ccguid = cp.Utils.CreateGuid()
            result.save(cp)
            Return result
        End Function
        '
        '====================================================================================================
        '
        Public Function Clone() As Object Implements ICloneable.Clone
            Return Me.MemberwiseClone()
        End Function
        ''' <summary>
        ''' Return a list of folders
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="FolderID">The id of the folder</param>
        ''' <returns></returns>
        Public Shared Function AllowFolderAccess(cp As CPBaseClass, FolderID As Integer, ParentID As Integer) As List(Of LibraryFolderModel)
            Dim result As New List(Of LibraryFolderModel)
            Try
                Dim SQL = "select top 1 *" _
                      & " from ccMemberRules M,ccLibraryFolderRules R" _
                      & " where M.MemberID=" & cp.User.Id _
                      & " and R.FolderID=" & FolderID _
                      & " and M.GroupID=R.GroupID" _
                      & " and R.Active<>0" _
                      & " and M.Active<>0" _
                      & " and ((M.DateExpires is null)or(M.DateExpires>" & cp.Db.EncodeSQLDate(Now) & "))"
                result = createList(cp, "(FolderID in (" & SQL & "))")
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        ''' <summary>
        ''' Return a list of folders
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <returns></returns>
        Public Shared Function LoadFolders_returnTopFolderId(cp As CPBaseClass, topFolderPath As String) As List(Of LibraryFolderModel)
            Dim result As New List(Of LibraryFolderModel)
            Try
                Dim SQL As String = "select Distinct" _
                    & " F.ID" _
                    & " ,F.ParentID" _
                    & " ,F.Name" _
                    & " ,(select top 1 ID from ccMemberRules where ccMemberRules.MemberID=" & cp.User.Id & " and ccMemberRules.GroupID=FR.GroupID) as Allowed" _
                    & " from (cclibraryfolders F left join ccLibraryFolderRules FR on FR.FolderID=F.ID)" _
                    & " where (f.active<>0)" _
                    & " order by f.name"
                result = createList(cp, "(FolderID in (" & SQL & "))")
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function

    End Class
End Namespace
