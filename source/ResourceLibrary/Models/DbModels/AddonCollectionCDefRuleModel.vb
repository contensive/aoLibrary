﻿
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Models     '<------ set namespace
    Public Class AddonCollectionCDefRuleModel        '<------ set set model Name and everywhere that matches this string
        Inherits baseModel
        Implements ICloneable
        '
        '====================================================================================================
        '-- const
        Public Const contentName As String = "Add-on Collection CDef Rules"      '<------ set content name
        Public Const contentTableName As String = "ccAddonCollectionCDefRules"   '<------ set to tablename for the primary content (used for cache names)
        Private Shadows Const contentDataSource As String = "default"             '<------ set to datasource if not default
        '
        '====================================================================================================
        ' -- instance properties
        Public Property CollectionID As Integer
        Public Property ContentID As Integer
        '
        '====================================================================================================
        Public Overloads Shared Function add(cp As CPBaseClass) As AddonCollectionCDefRuleModel
            Return add(Of AddonCollectionCDefRuleModel)(cp)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function create(cp As CPBaseClass, recordId As Integer) As AddonCollectionCDefRuleModel
            Return create(Of AddonCollectionCDefRuleModel)(cp, recordId)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function create(cp As CPBaseClass, recordGuid As String) As AddonCollectionCDefRuleModel
            Return create(Of AddonCollectionCDefRuleModel)(cp, recordGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function createByName(cp As CPBaseClass, recordName As String) As AddonCollectionCDefRuleModel
            Return createByName(Of AddonCollectionCDefRuleModel)(cp, recordName)
        End Function
        '
        '====================================================================================================
        Public Overloads Sub save(cp As CPBaseClass)
            MyBase.save(Of AddonCollectionCDefRuleModel)(cp)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Sub delete(cp As CPBaseClass, recordId As Integer)
            delete(Of AddonCollectionCDefRuleModel)(cp, recordId)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Sub delete(cp As CPBaseClass, ccGuid As String)
            delete(Of AddonCollectionCDefRuleModel)(cp, ccGuid)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Function createList(cp As CPBaseClass, sqlCriteria As String, Optional sqlOrderBy As String = "id") As List(Of AddonCollectionCDefRuleModel)
            Return createList(Of AddonCollectionCDefRuleModel)(cp, sqlCriteria, sqlOrderBy)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordName(cp As CPBaseClass, recordId As Integer) As String
            Return baseModel.getRecordName(Of AddonCollectionCDefRuleModel)(cp, recordId)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordName(cp As CPBaseClass, ccGuid As String) As String
            Return baseModel.getRecordName(Of AddonCollectionCDefRuleModel)(cp, ccGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordId(cp As CPBaseClass, ccGuid As String) As Integer
            Return baseModel.getRecordId(Of AddonCollectionCDefRuleModel)(cp, ccGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getCount(cp As CPBaseClass, sqlCriteria As String) As Integer
            Return baseModel.getCount(Of AddonCollectionCDefRuleModel)(cp, sqlCriteria)
        End Function
        '
        '====================================================================================================
        Public Overloads Function getUploadPath(fieldName As String) As String
            Return MyBase.getUploadPath(Of AddonCollectionCDefRuleModel)(fieldName)
        End Function
        '
        '====================================================================================================
        '
        Public Function Clone(cp As CPBaseClass) As AddonCollectionCDefRuleModel
            Dim result As AddonCollectionCDefRuleModel = DirectCast(Me.Clone(), AddonCollectionCDefRuleModel)
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

    End Class
End Namespace
