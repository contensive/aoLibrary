
Option Strict On
Option Explicit On

Imports Contensive.Addons.ResourceLibrary.Controllers
Imports Contensive.BaseClasses
Imports Contensive.Addons.ResourceLibrary.Controllers.genericController
Imports Contensive.Addons.ResourceLibrary.Models
Imports FastIndex5Class = Contensive.Addons.ResourceLibrary.Controllers.FastIndexClass

Namespace Contensive.Addons.ResourceLibrary
    '
    Public Class menuTreeClass
        '
        Private cp As CPBaseClass

        Friend Sub New(cp As CPBaseClass)
            Me.cp = cp
            EntryIndexName = New FastIndex5Class
            Randomize()
            MenuFlyoutNamePrefix = "id" & CStr(Int(9999 * Rnd()))
        End Sub
        '
        '==============================================================================
        '
        '   Creates custom menus
        '   Stores caches of the menus
        '   Stores the menu data, and can generate different kind
        '
        '==============================================================================
        '
        Const MenuStyleRollOverFlyout = 1
        Const MenuStyleTree = 2
        Const MenuStyleTreeList = 3
        Const MenuStyleFlyoutDown = 4
        Const MenuStyleFlyoutRight = 5
        Const MenuStyleFlyoutUp = 6
        Const MenuStyleFlyoutLeft = 7
        Const MenuStyleHoverDown = 8
        Const MenuStyleHoverRight = 9
        Const MenuStyleHoverUp = 10
        Const MenuStyleHoverLeft = 11
        '
        ' ----- Each menu item has an MenuEntry
        '
        Public Structure MenuEntryType
            Public Caption As String           ' What is displayed for this entry (does not need to be unique)
            Public Name As String              ' Unique name for this entry
            Public ParentName As String        ' Unique name of the parent entry
            Public Link As String              ' URL
            Public Image As String             ' Image
            Public ImageOver As String         ' Image Over
            Public ImageOpen As String         ' Image when menu is open
            Public NewWindow As Boolean        ' True opens link in a new window
            Public OnClick As String           ' Holds action for onClick
        End Structure

        Public iMenuFilePath As String
        '
        ' ----- Menu Entry storage
        '
        Public iEntryCount As Integer           ' Count of Menus in the object
        Public iEntrySize As Integer
        Public iEntry() As MenuEntryType
        '
        ' Private iDQMCount as integer            ' Count of Default Menus for this instance
        ' Private iDQMCLosed As Boolean       ' true if the menu has been closed
        '
        Public iTreeCount As Integer           ' Count of Tree Menus for this instance
        Public iMenuCloseString As String  ' String returned for closing menus
        '
        Public UsedEntries As String       ' String of EntryNames that have been used (for unique test)
        Public EntryIndexName As FastIndexClass
        Public EntryIndexID As FastIndexClass
        '
        ' ----- RollOverFlyout storage
        '
        Public MenuFlyoutNamePrefix As String    ' Random prefix added to element IDs to avoid namespace collision
        Public MenuFlyoutIcon_Local As String      ' string used to mark a button that has a non-hover flyout
        Public Const newmode = True
        '
        '===============================================================================
        '   Create a new Menu Entry
        '===============================================================================
        '
        Public Sub AddEntry(EntryName As String, ParentiEntryName As String, ImageLink As String, ImageOverLink As String, Link As String, Caption As String, Optional OnClickJavascript As String = "", Optional Ignore1 As String = "", Optional ImageOpenLink As String = "", Optional NewWindow As Boolean = False)
            Dim MenuEntrySize As Integer
            Dim iEntryName As String
            Dim UcaseEntryName As String
            Dim iNewWindow As Boolean
            '
            iEntryName = Replace(KmaEncodeMissingText(EntryName, ""), ",", " ")
            UcaseEntryName = UCase(iEntryName)
            '
            If (iEntryName <> "") And (InStr(1, UsedEntries & ",", "," & UcaseEntryName & ",", vbBinaryCompare) = 0) Then
                UsedEntries = UsedEntries & "," & UcaseEntryName
                If iEntryCount >= iEntrySize Then
                    iEntrySize = iEntrySize + 10
                    ReDim Preserve iEntry(iEntrySize)
                End If
                With iEntry(iEntryCount)
                    .Link = KmaEncodeMissingText(Link, "")
                    .Image = KmaEncodeMissingText(ImageLink, "")
                    .OnClick = KmaEncodeMissingText(OnClickJavascript, "")
                    If .Image = "" Then
                        '
                        ' No image, must have a caption
                        '
                        .Caption = KmaEncodeMissingText(Caption, iEntryName)
                    Else
                        '
                        ' Image present, caption is extra
                        '
                        .Caption = KmaEncodeMissingText(Caption, "")
                    End If
                    .Name = UcaseEntryName
                    .ParentName = UCase(KmaEncodeMissingText(ParentiEntryName, ""))
                    .ImageOver = KmaEncodeMissingText(ImageOverLink, "")
                    .ImageOpen = KmaEncodeMissingText(ImageOpenLink, "")
                    .NewWindow = KmaEncodeMissingBoolean(NewWindow, False)
                End With
                Call EntryIndexName.SetPointer(UcaseEntryName, iEntryCount)
                iEntryCount = iEntryCount + 1
            End If
        End Sub
        '
        '===============================================================================
        '   Returns the menu specified, if it is in local storage
        '
        '   It also creates the menu data in a close string that is returned in GetMenuClose.
        '   It must be done there so the link buttons height can be calculated.
        '===============================================================================
        '
        Public Function GetMenu(MenuName As String, Optional StyleSheetPrefix As String = "") As String
            GetMenu = GetTree(MenuName, "", KmaEncodeMissingText(StyleSheetPrefix, "ccTree"))
            'Exit Function
            ''
            'Dim Link As String
            'Dim EntryPointer as integer 
            'Dim UcaseMenuName As String
            'Dim LocalStyleSheetPrefix As String
            ''
            '' ----- Get the menu pointer
            ''
            'If iEntryCount > 0 Then
            '    UcaseMenuName = MenuName
            '    LocalStyleSheetPrefix = KmaEncodeMissingText(StyleSheetPrefix, "ccTree")
            '    For EntryPointer = 0 To iEntryCount - 1
            '        If iEntry(EntryPointer).Name = UcaseMenuName Then
            '            Exit For
            '        End If
            '    Next
            '    If EntryPointer < iEntryCount Then
            '        '
            '        ' ----- Build the linked -button-
            '        '
            '        Link = iEntry(EntryPointer).Link
            '        If Link = "" Then
            '            Link = "javascript: ;"
            '        End If
            '        '
            '        GetMenu = vbCrLf _
            '    & "<DIV id=""tree"" class=""" & LocalStyleSheetPrefix & "Root"" ></DIV>" & vbCrLf
            '        '
            '        '   Find the Menu Entry, and create the top element here
            '        '
            '        For EntryPointer = 0 To iEntryCount - 1
            '            With iEntry(EntryPointer)
            '                If .Name = UcaseMenuName Then
            '                    'iMenuCloseString = iMenuCloseString
            '                    GetMenu = GetMenu _
            '                & "<SCRIPT Language=""JavaScript"" type=""text/javascript"">" & vbCrLf _
            '                & "var DivLeft,DivTop,ElementObject; " & vbCrLf _
            '                & "DivTop = -18; " & vbCrLf _
            '                & "DivLeft = 0; " & vbCrLf _
            '                & "for (ElementObject=tree;  ElementObject.tagName!='BODY'; ElementObject = ElementObject.offsetParent) { " & vbCrLf _
            '                & "    DivTop = DivTop+ElementObject.offsetTop; " & vbCrLf _
            '                & "    DivLeft = DivLeft+ElementObject.offsetLeft; " & vbCrLf _
            '                & "    } " & vbCrLf _
            '                & "var menuBase = new  menuObject(DivTop,DivLeft); " & vbCrLf _
            '                & "menuBase.s[0] = new so(0,'" & .Caption & "','" & .Link & "','_blank',''); " & vbCrLf _
            '                & GetMenuTreeBranch(.Name, "menuBase.s[0]", "," & EntryPointer) _
            '                & "</SCRIPT>" & vbCrLf
            '                    ' & "<SCRIPT LANGUAGE=""JavaScript"" src=""/cclib/ClientSide/tree30.js""></SCRIPT>" & vbCrLf
            '                    Exit For
            '                End If
            '            End With
            '        Next
            '        '
            '        ' ----- Add what is needed to the close string, be carefull of the order
            '        '
            '        '
            '        ' increment the menu count
            '        '
            '        iTreeCount = iTreeCount + 1
            '    End If
            'End If
        End Function
        '
        '===============================================================================
        '   Gets the Menu Branch for the Tree Menu
        '===============================================================================
        '
        Private Function GetMenuTreeBranch(ParentName As String, JSObject As String, UsedEntries As String) As String
            '
            Dim EntryPointer As Integer
            Dim iUsedEntries As String
            Dim JSChildObject As String
            Dim SubMenuCount As Integer
            '
            iUsedEntries = UsedEntries
            SubMenuCount = 0
            For EntryPointer = 0 To iEntryCount - 1
                With iEntry(EntryPointer)
                    If .ParentName = ParentName Then
                        If (InStr(1, iUsedEntries & ",", "," & EntryPointer & ",") = 0) Then
                            JSChildObject = JSObject & ".s[" & SubMenuCount & "]"
                            iUsedEntries = iUsedEntries & "," & EntryPointer
                            GetMenuTreeBranch = GetMenuTreeBranch _
                        & JSChildObject & " = new so(0,'" & .Caption & "','" & .Link & "','_blank',''); " & vbCrLf _
                        & GetMenuTreeBranch(.Name, JSChildObject, iUsedEntries)
                            SubMenuCount = SubMenuCount + 1
                        End If
                    End If
                End With
            Next
        End Function
        '
        '===============================================================================
        '   Returns the menu specified, if it is in local storage
        '
        '   It also creates the menu data in a close string that is returned in GetMenuClose.
        '   It must be done there so the link buttons height can be calculated.
        '   Uses a simple UL/Stylesheet method, returning to the server with every click
        '===============================================================================
        '
        Private Function GetMenuTreeList(MenuName As String, OpenNodesList As String) As String
            '
            Dim EntryPointer As Integer
            Dim UcaseMenuName As String
            '
            ' ----- Get the menu pointer
            '
            If iEntryCount > 0 Then
                UcaseMenuName = UCase(MenuName)
                EntryPointer = EntryIndexName.GetPointer(UcaseMenuName)
                GetMenuTreeList = GetMenuTreeListBranch2(EntryPointer, "", OpenNodesList)
                Exit Function
            End If
        End Function
        '
        '===============================================================================
        '   Gets the Menu Branch for the Tree Menu
        '===============================================================================
        '
        Private Function GetMenuTreeListBranch2(NodePointer As Integer, UsedEntriesList As String, OpenNodesList As String) As String
            '
            Dim Link As String
            Dim EntryPointer As Integer
            Dim UcaseNodeName As String
            Dim Image As String
            Dim Caption As String
            '
            If iEntryCount > 0 Then
                '
                ' Output this node
                '
                If InStr(1, "," & CStr(NodePointer) & ",", "," & UsedEntriesList & ",") = 0 Then
                    GetMenuTreeListBranch2 = GetMenuTreeListBranch2 & "<ul Style=""list-style-type: none; margin-left: 20px"">"
                    '
                    ' The Node has not already been used in this branch
                    '
                    Caption = iEntry(NodePointer).Caption
                    Link = kmaEncodeHTML(cp, iEntry(NodePointer).Link)
                    If Link <> "" Then
                        Caption = "<A TARGET=""_blank"" HREF=""" & Link & """>" & Caption & "</A>"
                    End If
                    '
                    If InStr(1, "," & OpenNodesList & ",", "," & CStr(NodePointer) & ",") = 0 Then
                        '
                        ' The branch is closed
                        '
                        Image = iEntry(NodePointer).Image
                        GetMenuTreeListBranch2 = GetMenuTreeListBranch2 & "<li><A HREF=""?OpenNodesList=" & OpenNodesList & "&OpenNode=" & NodePointer & """><IMG SRC=""" & Image & """ HEIGHT=""18"" WIDTH=""18"" BORDER=0 ALT=""Open Folder"" /></A>&nbsp;" & Caption & "</li>"
                    Else
                        '
                        ' The branch is open
                        '
                        Image = iEntry(NodePointer).ImageOpen
                        If Image = "" Then
                            Image = iEntry(NodePointer).Image
                        End If
                        GetMenuTreeListBranch2 = GetMenuTreeListBranch2 _
                    & "<li>" _
                    & "<A HREF=""?OpenNodesList=" & OpenNodesList & "&CloseNode=" & NodePointer & """>" _
                    & "<IMG SRC=""" & Image & """ HEIGHT=""18"" WIDTH=""18"" BORDER=0 ALT=""Close Folder"" />" _
                    & "</A>&nbsp;" & Caption & "</li>"
                        '
                        ' Now output any child branches of this node
                        '
                        UcaseNodeName = UCase(iEntry(NodePointer).Name)
                        For EntryPointer = 0 To iEntryCount - 1
                            If (iEntry(EntryPointer).ParentName = UcaseNodeName) Then
                                GetMenuTreeListBranch2 = GetMenuTreeListBranch2 & GetMenuTreeListBranch2(EntryPointer, UsedEntriesList & "," & NodePointer, OpenNodesList)
                            End If
                        Next
                        ' GetMenuTreeListBranch2 = GetMenuTreeListBranch2 & GetMenuTreeListBranch2(iEntry(NodePointer).Name, UsedEntriesList & "," & CStr(NodePointer), OpenNodesList)
                    End If
                    GetMenuTreeListBranch2 = GetMenuTreeListBranch2 & "</ul>" & vbCrLf
                End If
            End If
        End Function
        '
        '===============================================================================
        '   Returns the menu specified, if it is in local storage
        '
        '   It also creates the menu data in a close string that is returned in GetTreeClose.
        '   It must be done there so the link buttons height can be calculated.
        '===============================================================================
        '
        Public Function GetTree(MenuName As String, OpenMenuName As String, Optional StyleSheetPrefix As String = "") As String
            GetTree = ""
            Dim Link As String
            Dim EntryPointer As Integer
            Dim UcaseMenuName As String
            Dim UsedEntries As String
            Dim Caption As String
            Dim JSString As String
            '
            ' ----- Get the menu pointer
            '
            If iEntryCount > 0 Then
                UcaseMenuName = UCase(MenuName)
                If StyleSheetPrefix = "" Then
                    StyleSheetPrefix = "ccTree"
                End If
                If True Then
                    '
                    ' ----- Build the linked -button-
                    '
                    Link = iEntry(EntryPointer).Link
                    If Link = "" Then
                        Link = "javascript: ;"
                    End If
                    '
                    '   Find the Menu Entry, and create the top element here
                    '
                    UsedEntries = ""
                    For EntryPointer = 0 To iEntryCount - 1
                        With iEntry(EntryPointer)
                            If .Name = UcaseMenuName Then
                                Caption = .Caption
                                If .Link <> "" Then
                                    Caption = "<a href=""" & kmaEncodeHTML(cp, .Link) & """>" & Caption & "</a>"
                                End If
                                UsedEntries = UsedEntries & "," & CStr(EntryPointer)
                                GetTree = "" _
                            & vbCrLf & "<ul class=mktree id=tree" & iTreeCount & ">" & vbCrLf _
                            & vbCrLf & " <li id=""" & .Name & """><span class=mkc>" & Caption & "</span>" _
                            & vbCrLf & " <ul>" & vbCrLf _
                            & GetMKTreeBranch(UcaseMenuName, UsedEntries, 2) _
                            & vbCrLf & " </ul>" & vbCrLf _
                            & vbCrLf & "</li></ul>" & vbCrLf
                                Exit For
                            End If
                        End With
                    Next
                    If UsedEntries = "" Then
                        GetTree = "" _
                            & vbCrLf & "<ul class=mktree id=tree" & iTreeCount & ">" _
                            & GetMKTreeBranch(UcaseMenuName, UsedEntries, 1) _
                            & vbCrLf & "</ul>" & vbCrLf
                    End If
                    GetTree &= "<script src=/resourcelibrary/mktree.js></script>"
                    GetTree &= "<script type=""text/javascript"">convertTrees();"
                    If OpenMenuName <> "" Then
                        JSString = UCase(OpenMenuName)
                        JSString = Replace(JSString, "\", "\\")
                        JSString = Replace(JSString, vbCrLf, "\n")
                        JSString = Replace(JSString, "'", "\'")
                        'Call Main.AddOnLoadJavascript("expandToItem('tree" & iTreeCount & "','" & JSString & "');")
                        GetTree = GetTree & "expandToItem('tree" & iTreeCount & "','" & JSString & "');"
                    End If
                    GetTree = GetTree & "</script>"
                    '
                    ' increment the menu count
                    '
                    iTreeCount = iTreeCount + 1
                End If
            End If
        End Function
        '
        '===============================================================================
        '   Gets the Menu Branch for the Tree Menu
        '===============================================================================
        '
        Private Function GetMKTreeBranch(ParentName As String, UsedEntries As String, Depth As Integer) As String
            '
            Dim EntryPointer As Integer
            Dim iUsedEntries As String
            Dim SubMenuCount As Integer
            Dim ChildMenu As String
            Dim Caption As String
            '
            iUsedEntries = UsedEntries
            SubMenuCount = 0
            For EntryPointer = 0 To iEntryCount - 1
                With iEntry(EntryPointer)
                    If .ParentName = ParentName Then
                        If (InStr(1, iUsedEntries & ",", "," & EntryPointer & ",") = 0) Then
                            Caption = .Caption
                            If .OnClick <> "" And .Link <> "" Then
                                Caption = "<a href=""" & kmaEncodeHTML(cp, .Link) & """ onClick=""" & .OnClick & """>" & Caption & "</a>"
                            ElseIf .OnClick <> "" Then
                                Caption = "<a href=""#"" onClick=""" & .OnClick & """>" & Caption & "</a>"
                            ElseIf .Link <> "" Then
                                Caption = "<a href=""" & kmaEncodeHTML(cp, .Link) & """>" & Caption & "</a>"
                            Else
                                Caption = Caption
                            End If
                            iUsedEntries = iUsedEntries & "," & EntryPointer

                            ChildMenu = GetMKTreeBranch(.Name, iUsedEntries, Depth + 1)
                            If newmode Then
                                If ChildMenu = "" Then
                                    GetMKTreeBranch = GetMKTreeBranch _
                                & vbCrLf & "<li class=mklb id=""" & .Name & """ >" _
                                & "<div class=""mkd"">" _
                                & "<span class=mkb>&nbsp;</span>" _
                                & "</div>" _
                                & Caption _
                                & "</li>"
                                Else
                                    '
                                    ' 3/18/2010 changes to keep firefox from blocking clicks
                                    '
                                    GetMKTreeBranch = GetMKTreeBranch _
                                & vbCrLf & "<li class=""mklc"" id=""" & .Name & """ >" _
                                & "<div class=""mkd"" >" _
                                & "<span class=mkb onclick=""mkClick(this)"">&nbsp;</span>" _
                                & "</div>" _
                                & Caption _
                                & vbCrLf & "<ul>" _
                                & ChildMenu _
                                & vbCrLf & "</ul>" _
                                & "</li>"
                                End If
                            Else
                                If ChildMenu <> "" Then
                                    ChildMenu = "" _
                                & vbCrLf & "<ul>" _
                                & ChildMenu _
                                & vbCrLf & "</ul>" _
                                & ""
                                End If
                                GetMKTreeBranch = GetMKTreeBranch _
                            & vbCrLf & "<li class=mklc id=""" & .Name & """>" _
                            & Caption _
                            & ChildMenu _
                            & vbCrLf & "</li>"
                            End If
                            SubMenuCount = SubMenuCount + 1
                        End If
                    End If
                End With
            Next
        End Function
    End Class
End Namespace
