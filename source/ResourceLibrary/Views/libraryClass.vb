
'Option Strict On
Option Explicit On

Imports System.Net
Imports System.Web
Imports Contensive.Addons.ResourceLibrary.Controllers
Imports Contensive.BaseClasses
Imports Contensive.Addons.ResourceLibrary.Controllers.genericController
Imports Contensive.Addons.ResourceLibrary.Models
Imports Contensive.VbConversion.Contensive.VbConversion
Imports Contensive.VbConversion

Namespace Contensive.Addons.ResourceLibrary.Views
    '
    Public Class libraryClass
        Inherits AddonBaseClass
        '
        'Private main As Contensive.vbConversion.MainClass
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                '
                returnHtml = GetContent(CP)
                '
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
            End Try
            Return returnHtml
        End Function
        '
        '
        Public Class FileType
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
        Public iconFiles As New List(Of FileType)
        'Public IconFiles() As FileType
        'Public IconFileCnt As Integer
        '
        '
        '
        Public Class FolderType
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
        Public folders() As FolderType = {}
        Public folderCnt As Integer
        Public FolderIdIndex As New VbConversion.FastIndexClass
        Public FolderNameIndex As New VbConversion.FastIndexClass
        Public FolderPathIndex As New VbConversion.FastIndexClass
        '
        Public FolderSelect As String
        '
        ' -----------------------------------------------------------------------------------
        ' ----- Publics
        ' -----------------------------------------------------------------------------------
        ' ----- not used
        '
        Public UserMemberID As Integer
        Public RequestStream As String
        '
        ' ----- Icons used
        '
        Public Const IconFolderOpen = "<img src=""/ResourceLibrary/IconFolderOpen.gif"" border=""0"" width=""22"" height=""23"" ALT=""Close this folder"">"
        Public Const IconFolderClosed = "<img src=""/ResourceLibrary/IconFolderClosed.gif"" border=""0"" width=""22"" height=""23"" alt=""Open this folder"">"
        Public Const IconFolderAdd = "<img src=""/ResourceLibrary/IconFolderAdd2.gif"" border=""0"" width=""22"" height=""23"" alt=""Add a New folder"">"
        Public Const IconFolderEdit = "<img src=""/ResourceLibrary/IconFolderEdit.gif"" border=""0"" width=""22"" height=""23"" alt=""Edit this folder"">"
        Public Const IconFile = "<img src=""/ResourceLibrary/IconFile.gif"" border=""0"" width=""22"" height=""23"" alt=""file"">"
        Public Const IconFileAdd = "<img src=""/ResourceLibrary/IconContentAdd.gif"" border=""0"" width=""18"" height=""22"" alt=""Add a New  file"">"
        Public Const IconFileEdit = "<img src=""/ResourceLibrary/IconContentEdit.gif"" border=""0"" width=""18"" height=""22"" alt=""Edit this file"">"
        Public Const IconPreview = "<img src=""/ResourceLibrary/IconPreview.gif"" border=""0"" width=""22"" height=""23"" alt=""Preview this image"">"
        Public Const IconDownload = "<img src=""/ResourceLibrary/IconDownload3.gif"" border=""0"" width=""16"" height=""16"" alt=""Select this download"" valign=""absmiddle"">"
        Public Const IconCreateImage = "<img src=""/ResourceLibrary/IconimagePlace.gif"" border=""0"" width=""18"" height=""22"" alt=""Select this image"">"
        Public Const IconCreateDownload = "<img src=""/ResourceLibrary/IconDownload3.gif"" border=""0"" width=""16"" height=""16"" alt=""Select this download"" valign=""absmiddle"">"
        Public Const IconSpacer = "<img src=""/ResourceLibrary/spacer.gif"" width=""22"" height=""23"">"
        Public Const IconView = "<img src=""/ResourceLibrary/IconView.gif"" border=""0"" width=""22"" height=""23"" alt=""Preview this file"">"
        Public Const IconImage = "<img src=""/ResourceLibrary/IconImage2.gif"" border=""0"" width=""22"" height=""23"" alt=""Image"">"
        Public Const IconPDF = "<img src=""/ResourceLibrary/IconPDF.gif"" border=""0"" width=""16"" height=""16"" alt=""Adobe Pdf"">"
        Public Const IconOther = "<img src=""/ResourceLibrary/IconFile.gif"" border=""0"" width=""22"" height=""23"" alt=""Unrecognized File Type"">"
        Public Const IconNoFile = "<img src=/ResourceLibrary/BulletRound2.gif width=5 height=5>"
        '
        ' ----- SelectResource Support
        '       This means the resource library supports buttons that allow objects to be
        '       placed on different page from the resource library, like an Editor
        '
        Public AllowPlace As Boolean
        '
        ' ----- If an editor is used to call the resource library, the window.opener.insertresource()
        '       call needs the object name of the editor so the contents can be copied to the invisible
        '       form field after the changes (no onchange event available)
        '
        Public SelectResourceEditorObjectName As String
        '
        ' ----- If AllowPlace is true and SelectLinkObjectName<>"", the RL is being used as a link selector
        '       When the 'place' icon is clicked, the URL of the resource is copied to the window.opener.[selectlinkobjectname]
        '
        Public SelectLinkObjectName As String
        '
        ' ----- Blocks the folder list in the left hand side
        '
        Public blockFolderNavigation As Boolean
        '
        ' -----------------------------------------------------------------------------------
        ' ----- Privates
        ' -----------------------------------------------------------------------------------
        '
        Public iMinRows As Integer
        Public iFolderID As Integer                      ' Current Folder being Displayed, 0 for root
        Public SourceMode As Integer                      '
        '
        '        ' SourceMode
        '        '   3/6/2010 - moved codes up to capture the 0 case and it to page
        '        '   1 = From Editor Object or Link selector: allow image and download insert, provide close button
        '        '   2 = From Editor Image Properties: allow image insert, provide close button
        '        '   3 = From Admin site, no inserts, and provide cancel button
        'Const SourceModeOnPage = 1
        'Const SourceModeFromDownloadRequest = 2
        'Const SourceModeFromLinkDialog = 3
        '   0 = From Editor Object selector: allow image and download insert, provide close button
        '   1 = From Editor Image Properties: allow image insert, provide close button
        '   2 = From Admin site, no inserts, and provide cancel button
        Public Const SourceModeFromDownloadRequest = 0
        Public Const SourceModeFromLinkDialog = 1
        Public Const SourceModeOnPage = 2
        '
        '   0 caller is the editor directly, clicking on action icons calls InsertImaage, etc
        '   1 caller is the editor image page, clicking on action icons calls the image page methods
        '
        Public HoldPosition As Integer
        'Private main As MainClass

        '
        ''=====================================================================================
        '''' <summary>
        '''' AddonDescription
        '''' </summary>
        '''' <param name="CP"></param>
        '''' <returns></returns>
        'Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
        '    Dim result As String = ""
        '    'Dim sw As New Stopwatch : sw.Start()
        '    Try
        '        '
        '        ' -- initialize application. If authentication needed and not login page, pass true
        '        Using ae As New applicationController(CP, False)
        '            main = New vbConversion.Contensive.VbConversion.MainClass(CP)
        '            '
        '            ' -- your code
        '            result = GetContent(CP)
        '            If ae.packageErrorList.Count > 0 Then
        '                result = "Hey user, this happened - " & Join(ae.packageErrorList.ToArray, "<br>")
        '            End If
        '        End Using
        '    Catch ex As Exception
        '        CP.Site.ErrorReport(ex)
        '    End Try
        '    Return result
        'End Function

        '
        '=================================================================================
        '   Aggregate Object Interface
        '=================================================================================
        '
        Public Function GetContent(cp As CPBaseClass) As String
            Dim result As String = ""
            Try
                Dim topFolderPath As String = ""
                Dim AllowGroupAdd As Boolean
                Dim OptionString As String = ""
                '
                topFolderPath = cp.Doc.GetText("RootFolderName")
                Call cp.Site.TestPoint("topFolderPath=[" & topFolderPath & "]")
                AllowGroupAdd = cp.Utils.EncodeBoolean(cp.Doc.GetBoolean("AllowGroupAdd"))
                AllowPlace = cp.Utils.EncodeBoolean(cp.Doc.GetBoolean("AllowSelectResource"))
                SelectResourceEditorObjectName = cp.Doc.GetText("SelectResourceEditorObjectName")
                SelectLinkObjectName = cp.Doc.GetText("SelectLinkObjectName")
                blockFolderNavigation = cp.Utils.EncodeBoolean(cp.Doc.GetBoolean("Block Folder Navigation"))
                '
                ' topFolder should be in this format toptier\tier2\tier2
                '   all lowercase, no leading or trailing slashes, backslashs, remove 'root\'
                '
                topFolderPath = Trim(topFolderPath)
                topFolderPath = LCase(topFolderPath)
                topFolderPath = Replace(topFolderPath, "/", "\")
                If Left(topFolderPath, 4) = "root" Then
                    topFolderPath = Mid(topFolderPath, 5)
                End If
                If Left(topFolderPath, 1) = "\" Then
                    topFolderPath = Mid(topFolderPath, 2)
                End If
                If Right(topFolderPath, 1) = "\" Then
                    topFolderPath = Mid(topFolderPath, 1, Len(topFolderPath) - 1)
                End If
                '
                GetContent = GetForm(cp, topFolderPath, AllowGroupAdd)
                result = GetContent
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function

        '
        '=================================================================================
        ' Returns the Resource Library HTML.
        '   This HTML does not include the HTML, HEAD or BODY tags.
        '=================================================================================
        '
        Private Function GetForm(cp As CPBaseClass, topFolderPath As String, AllowGroupAdd As Boolean) As String
            Dim result As String = ""
            Try
                Const LibraryFileTypespathFilename = "ResourceLibrary\LibraryConfig.xml"
                Dim FolderAccess As Boolean
                Dim hint As String
                '
                hint = "000"
                '
                Const Image5 = "<img src=/ResourceLibrary/spacer.gif width=5 height=1>"
                Const Image10 = "<img src=/ResourceLibrary/spacer.gif width=10 height=1>"
                Const Image15 = "<img src=/ResourceLibrary/spacer.gif width=15 height=1>"
                Const Image20 = "<img src=/ResourceLibrary/spacer.gif width=20 height=1>"
                Const Image30 = "<img src=/ResourceLibrary/spacer.gif width=30 height=1>"
                Const Image50 = "<img src=/ResourceLibrary/spacer.gif width=50 height=1>"
                '
                Dim ButtonBarStyle As String = "" _
                        & " color: black;" _
                        & " font-weight: bold;" _
                        & " padding: 5px;" _
                        & " background-color: #a0a0a0;" _
                        & " border-bottom: 1px solid #e0e0e0;" _
                        & " border-right: 1px solid #e0e0e0;" _
                        & " border-top: 1px solid #808080;" _
                        & " border-left: 1px solid #808080;"
                '
                Dim OptionPanelStyle As String = "" _
                        & " color: black;" _
                        & " font-weight: bold;" _
                        & " padding: 5px;" _
                        & " background-color: #d0d0d0;" _
                        & " border-bottom: 1px solid #e0e0e0;" _
                        & " border-right: 1px solid #e0e0e0;" _
                        & " border-top: 1px solid #a0a0a0;" _
                        & " border-left: 1px solid #a0a0a0;"
                '
                If Not (False) Then
                    '
                    ' Determine Current Folder
                    '
                    hint = "001"
                    'BuildVersion = cp.Site.GetText("build version")
                    Dim IsContentManagerFiles As Boolean = cp.User.IsContentManager("Library Files")
                    Dim IsContentManagerFolders As Boolean = cp.User.IsContentManager("Library Folders")
                    Dim Button As String = cp.Doc.GetText("Button")
                    Dim FileTypeFilter As String = LCase(cp.Doc.GetText("ffilter"))
                    Call cp.Doc.AddRefreshQueryString("ffilter", FileTypeFilter)
                    Dim AllowThumbnails As Boolean = cp.User.GetBoolean("LibraryAllowthumbnails", "0")
                    Dim FolderIDString As String = cp.Doc.GetText("folderid")
                    Dim currentFolderID As Integer = cp.Utils.EncodeInteger(FolderIDString)
                    If FolderIDString <> "" Then
                        Call cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString())
                    Else
                        currentFolderID = cp.User.GetInteger("Libraryfolderid", "0")
                    End If
                    '
                    ' Load Folder cache
                    '
                    hint = "010, topFolderPath=" & topFolderPath
                    Dim topFolderID As Integer = LoadFolders_returnTopFolderId(cp, topFolderPath)
                    '
                    Dim reloadFolderCache As Boolean = False
                    Dim currentFolderPtr As Integer
                    '
                    ' verify that current folder has viewAccess (if not jumpt to root)
                    '
                    If currentFolderID <> 0 Then
                        currentFolderPtr = FolderIdIndex.getPtr(currentFolderID)
                        If (currentFolderPtr > UBound(folders)) Or (currentFolderPtr < 0) Then
                            currentFolderPtr = 0
                        End If
                        If currentFolderID < 0 Then
                            currentFolderID = 0
                            Call cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString())
                        ElseIf Not folders(currentFolderPtr).hasViewAccess Then
                            currentFolderID = 0
                            Call cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString())
                        End If
                    End If
                    '
                    ' determine if current folder has modify access
                    '
                    hint = "020"
                    Dim currentFolderHasModifyAccess As Boolean = False
                    If (cp.User.IsAdmin Or IsContentManagerFiles Or IsContentManagerFolders) Then
                        '
                        ' you get modify access if you can modify the content
                        '
                        currentFolderHasModifyAccess = True
                    ElseIf currentFolderID = 0 Then
                        '
                        ' only admin and content managers of files and folders have modify access to root folder
                        '
                    Else
                        '
                        ' others have modify access to this folder if they are in a modify access group
                        '
                        currentFolderPtr = FolderIdIndex.getPtr(CStr(currentFolderID))
                        If currentFolderPtr >= 0 Then
                            currentFolderHasModifyAccess = folders(currentFolderPtr).hasModifyAccess
                        End If
                    End If
                    'topFolderID = GetFolderID(topFolderPath)
                    '
                    ' Load IconFiles
                    '
                    hint = "030"
                    Dim doc As Xml.XmlDocument = New Xml.XmlDocument
                    Dim FilePath As String = cp.Request.Protocol & cp.Request.Host & cp.Site.FilePath
                    Dim ConfigFilename As String = cp.Site.PhysicalFilePath & LibraryFileTypespathFilename
                    Call doc.Load(ConfigFilename)
                    Dim Ptr As Integer
                    If False Then
                        '
                        ' Error
                        '
                        'Call AppendLogFile2( "Server", "AddonInstallClass", "DownloadCollectionFiles, The GetCollection request for GUID [" & CollectionGuid & "] failed. The error was [" & doc.parseError.reason & "]")
                    Else
                        hint = "040"
                        If (LCase(doc.DocumentElement.Name) <> LCase("libraryconfig")) Then
                            'Return_ErrorMessage = "The collection file from the server was not valid for collection [" & CollectionGuid & "]"
                            'DownloadCollectionFiles = False
                            'Call AppendClassLogFile("Server", "AddonInstallClass", "DownloadCollectionFiles, The GetCollection request for GUID [" & CollectionGuid & "] named [" & Collectionname & "] returned a file with a bad format. The root node was [" & doc.documentElement.Name & "] but [" & DownloadFileRootNode & "] was expected.")
                        Else
                            If doc.DocumentElement.ChildNodes.Count = 0 Then
                                'Return_ErrorMessage = "The collection file from the server was empty for collection [" & CollectionGuid & "]"
                                'Call AppendClassLogFile("Server", "AddonInstallClass", "DownloadCollectionFiles, The GetCollection request for GUID [" & CollectionGuid & "] named [" & Collectionname & "] returned a file with no nodes. The collection was probably not found")
                                'DownloadCollectionFiles = False
                            Else
                                With doc.DocumentElement
                                    Ptr = 0
                                    hint = "050"
                                    Dim baseNode As Xml.XmlElement
                                    For Each baseNode In .ChildNodes
                                        hint = "060"
                                        Select Case LCase(baseNode.Name)
                                            Case "filetype"
                                                hint = "070"
                                                Ptr = Ptr + 1
                                                Dim newFileType As New FileType
                                                iconFiles.Add(newFileType)
                                                'Dim IconCnt As Integer
                                                'If Ptr >= IconCnt Then
                                                '    IconCnt = IconCnt + 10
                                                '    ReDim Preserve IconFiles(IconCnt)
                                                'End If
                                                With newFileType
                                                    Dim typeNode As Xml.XmlElement
                                                    For Each typeNode In baseNode.ChildNodes
                                                        Select Case LCase(typeNode.Name)
                                                            Case "name"
                                                                .Name = typeNode.Value
                                                            Case "filetypeid"
                                                                .FileTypeID = cp.Utils.EncodeInteger(typeNode.Value)
                                                            Case "extensionlist"
                                                                .ExtensionList = typeNode.Value
                                                            Case "isdownload"
                                                                .IsDownload = cp.Utils.EncodeBoolean(typeNode.Value)
                                                            Case "isimage"
                                                                .IsImage = cp.Utils.EncodeBoolean(typeNode.Value)
                                                            Case "isvideo"
                                                                .IsVideo = cp.Utils.EncodeBoolean(typeNode.Value)
                                                            Case "isflash"
                                                                .IsFlash = cp.Utils.EncodeBoolean(typeNode.Value)
                                                            Case "iconlink"
                                                                .IconFilename = typeNode.Value
                                                            Case "mediaiconlink"
                                                                .MediaIconFilename = typeNode.Value
                                                            Case "downloadiconlink"
                                                                .DownloadIconFilename = typeNode.Value
                                                        End Select
                                                    Next
                                                End With
                                        End Select
                                    Next
                                    'IconFileCnt = Ptr
                                End With
                            End If
                        End If
                    End If
                    '
                    ' Verify default icons
                    '
                    hint = "100"
                    Dim DefaultIcon As String = "\cclib\images\IconImage2.gif"
                    Dim DefaultMedia As String = "\cclib\images\Iconimage2Media.gif"
                    Dim DefaultDownload As String = "\cclib\images\Iconimage2Download.gif"
                    '
                    If cp.Doc.GetText("SourceMode") = "" Then
                        SourceMode = SourceModeOnPage
                    Else
                        SourceMode = cp.Doc.GetInteger("SourceMode")
                    End If
                    Call cp.Doc.AddRefreshQueryString("SourceMode", SourceMode.ToString())
                    '
                    ' ----- verify currentFolderID
                    '
                    If currentFolderID = 0 Then
                        '
                        ' No folder give, use root folder, no owner
                        currentFolderID = topFolderID
                        Call cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString())
                    End If
                    Dim folder As LibraryFolderModel = LibraryFolderModel.create(cp, currentFolderID)
                    'Dim FolderGroupName As String
                    Dim FolderParentID As Integer
                    If (folder IsNot Nothing) Then
                        FolderParentID = folder.ParentID
                    End If
                    If (topFolderID <> currentFolderID) And (topFolderID <> FolderParentID) Then
                        '
                        ' Check if Folder is under the given root folder
                        If Not IsInFolder(cp, topFolderID, currentFolderID) Then
                            '
                            ' Current folder is not in Root Folder, Use Root Folder
                            currentFolderID = topFolderID
                            Call cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString())
                        End If
                    End If
                    '
                    ' ----- Determine if select, place or edit icons are allowed
                    '
                    Dim ColumnCnt As Integer = 5
                    Dim AllowPlaceColumn As Boolean = AllowPlace And ((SourceMode = SourceModeFromLinkDialog) Or (SourceMode = SourceModeFromDownloadRequest))
                    If AllowPlaceColumn Then
                        ColumnCnt = ColumnCnt + 1
                    End If
                    Dim AllowEditColumn As Boolean = (IsContentManagerFiles Or IsContentManagerFolders)
                    If AllowEditColumn Then
                        ColumnCnt = ColumnCnt + 1
                    End If
                    Dim AllowSelectColumn As Boolean = currentFolderHasModifyAccess
                    If AllowSelectColumn Then
                        ColumnCnt = ColumnCnt + 1
                    End If
                    '
                    ' ----- Setup folder editing
                    Dim AllowFolderAuthoring As Boolean = IsContentManagerFolders
                    Dim FolderCID As Integer
                    If AllowFolderAuthoring Then
                        FolderCID = cp.Content.GetID("Library Folders")
                    End If
                    '
                    ' ----- Setup file editing
                    Dim AllowFileAuthoring As Boolean = IsContentManagerFiles
                    Dim FileCID As Integer
                    If AllowFileAuthoring Then
                        FileCID = cp.Content.GetID("Library Files")
                    End If
                    'Dim FolderGroupID as integer
                    '
                    ' ----- Setup Local File Management
                    '       Allow if Content Manager or user has group membership
                    '       Always allow, everyone has access to the root folder, then if you can get to the folder, let em upload
                    Dim AllowLocalFileAdd As Boolean = True
                    '
                    ' ----- Process input
                    Dim Pos As Integer
                    Dim AltSizeList As String
                    Dim FilenameNoExtension As String
                    Dim FileExtension As String
                    Dim FileTypeID As Integer
                    Dim RowCount As Integer
                    If Button <> "" Then
                        AllowThumbnails = cp.Doc.GetBoolean("AllowThumbnails")
                        Call cp.User.SetProperty("LibraryAllowthumbnails", AllowThumbnails.ToString())
                        Select Case Button
                            Case ButtonCancel
                                '
                                ' CAncel button, just redirect back to the current page
                                Call cp.Response.Redirect("#")
                            Case ButtonDelete
                                '
                                '
                                '
                                RowCount = cp.Doc.GetInteger("RowCount")
                                Dim DeleteFileID As Integer
                                Dim DeleteFolderID As Integer
                                If RowCount > 0 Then
                                    For Ptr = 0 To RowCount - 1
                                        If cp.Doc.GetBoolean("Row" & Ptr) Then
                                            DeleteFolderID = cp.Doc.GetInteger("Row" & Ptr & "FolderID")
                                            If DeleteFolderID <> 0 Then
                                                'Call Main.WriteStream("Deleting Folder " & FolderID)
                                                Call cp.Content.Delete("Library Folders", "id=" & DeleteFolderID)
                                                reloadFolderCache = True
                                            End If
                                            DeleteFileID = cp.Doc.GetInteger("Row" & Ptr & "FileID")
                                            If DeleteFileID <> 0 Then
                                                'Call Main.WriteStream("Deleting File " & FileID)
                                                Call cp.Content.Delete("Library Files", "id=" & DeleteFileID)
                                                reloadFolderCache = True
                                            End If
                                        End If
                                    Next
                                End If
                            Case ButtonApply
                                '
                                ' Move Files
                                '
                                If cp.Doc.GetBoolean("Move") Then
                                    Dim targetFolderId As Integer = cp.Doc.GetInteger("MoveFolderID")
                                    RowCount = cp.Doc.GetInteger("RowCount")
                                    If RowCount > 0 Then
                                        For Ptr = 0 To RowCount - 1
                                            If cp.Doc.GetBoolean("Row" & Ptr) Then
                                                Dim MoveFolderID As Integer = cp.Doc.GetInteger("Row" & Ptr & "FolderID")
                                                Dim MoveFileID As Integer = cp.Doc.GetInteger("Row" & Ptr & "FileID")
                                                If MoveFolderID <> 0 Then
                                                    Call cp.Db.ExecuteSQL("update ccLibraryFolders set ParentID=" & targetFolderId & " where ID=" & MoveFolderID)
                                                    reloadFolderCache = True
                                                ElseIf MoveFileID <> 0 Then
                                                    Call cp.Db.ExecuteSQL("update ccLibraryFiles set FolderID=" & targetFolderId & " where ID=" & MoveFileID)
                                                    reloadFolderCache = True
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                                '
                                ' Upload
                                '
                                If AllowLocalFileAdd Then
                                    '
                                    ' Add Folders
                                    '
                                    hint = "300"
                                    Dim AddFolderCount As Integer = cp.Doc.GetInteger("AddFolderCount")
                                    Dim UploadPointer As Integer
                                    For UploadPointer = 1 To AddFolderCount
                                        Dim folderName As String = cp.Doc.GetText("FolderName." & UploadPointer)
                                        If folderName <> "" Then
                                            If IsContentManagerFolders And (Not cp.User.IsAdmin) And (currentFolderID = 0) Then
                                                '
                                                ' Content Managers can not add folders to the root folder
                                                '
                                                Call cp.UserError.Add("Your account does not have access to add new folders to the root folder.")
                                                Exit For
                                            Else
                                                Dim libraryFolder As Models.LibraryFolderModel = Models.LibraryFolderModel.add(cp)
                                                libraryFolder.name = folderName
                                                libraryFolder.Description = cp.Doc.GetText("FolderDescription." & UploadPointer)
                                                libraryFolder.ParentID = currentFolderID
                                                libraryFolder.save(cp)
                                                'cS = Main.InsertCSRecord("Library Folders")
                                                'If Main.IsCSOK(cS) Then
                                                '    Copy = cp.Doc.GetText("FolderDescription." & UploadPointer)
                                                '    Call Main.SetCS(cS, "Name", folderName)
                                                '    Call Main.SetCS(cS, "Description", Copy)
                                                '    If currentFolderID <> 0 Then
                                                '        Call Main.SetCS(cS, "ParentID", currentFolderID)
                                                '    End If
                                                'End If
                                                'Call Main.closecs(cS)
                                                reloadFolderCache = True
                                            End If
                                        End If
                                    Next
                                    '
                                    ' Upload files
                                    '
                                    hint = "400"
                                    Dim UploadCount As Integer = cp.Doc.GetInteger("LibraryUploadCount")
                                    Dim ImageFilename As String = ""
                                    'Dim imagefileFolderId As Integer = cp.Doc.GetInteger("FolderID")
                                    For UploadPointer = 1 To UploadCount
                                        Dim imageRequestName As String = RequestNameLibraryUpload & "." & UploadPointer
                                        ImageFilename = cp.Doc.GetText(RequestNameLibraryUpload & "." & UploadPointer)
                                        If ImageFilename <> "" Then
                                            hint = "410"
                                            Dim libraryFile As LibraryFileModel = LibraryFileModel.add(cp)


                                            Dim libraryName As String = cp.Doc.GetText(RequestNameLibraryName & "." & UploadPointer)
                                            If libraryName = "" Then
                                                libraryName = ImageFilename
                                            End If
                                            libraryFile.name = libraryName
                                            Dim libraryDescription = cp.Doc.GetText(RequestNameLibraryDescription & "." & UploadPointer)
                                            If libraryDescription = "" Then
                                                libraryDescription = ImageFilename
                                            End If
                                            FileExtension = ""
                                            FilenameNoExtension = ""
                                            AltSizeList = ""
                                            Pos = InStrRev(ImageFilename, ".")
                                            If Pos > 0 Then
                                                FileExtension = Mid(ImageFilename, Pos + 1)
                                                FilenameNoExtension = Left(ImageFilename, Pos - 1)
                                            End If
                                            '''''libraryFile.Filename.upload(cp, imageRequestName)

                                            Dim VirtualFilePathPage As String = libraryFile.getUploadPath("filename")


                                            Dim VirtualFilePath As String = Replace(VirtualFilePathPage, ImageFilename, "")
                                            libraryFile.Description = libraryDescription
                                            libraryFile.FolderID = currentFolderID
                                            cp.Html.ProcessInputFile(imageRequestName, VirtualFilePath)

                                            libraryFile.FileSize = GetFileSize(cp, cp.Site.PhysicalFilePath & libraryFile.name)
                                            Dim FileType As String = ""
                                            hint = "425"
                                            FileTypeID = GetFileTypeID(cp, ImageFilename)
                                            libraryFile.FileTypeID = FileTypeID
                                            libraryFile.name = libraryName
                                            libraryFile.Description = libraryDescription
                                            libraryFile.Filename = VirtualFilePath & ImageFilename
                                            libraryFile.ModifiedDate = Date.Now
                                            libraryFile.save(cp)

                                            reloadFolderCache = True
                                        End If
                                    Next
                                End If

                        End Select
                    End If
                    hint = "500"
                    If reloadFolderCache Then
                        folderCnt = 0
                        topFolderID = LoadFolders_returnTopFolderId(cp, topFolderPath)
                        reloadFolderCache = False
                    End If
                    '
                    ' Housekeep potential issue where a parent is deleted and child does not show
                    '
                    Dim SQL As String = "update cclibraryfolders" _
                            & " Set parentid=null" _
                            & " where id in" _
                            & " (" _
                            & " select c.id from (cclibraryfolders c left join cclibraryfolders p on p.id=c.parentid)" _
                            & " where p.ID Is Null" _
                            & " and c.parentid is not null" _
                            & " and c.parentid<>0" _
                            & " )"
                    Call cp.Db.ExecuteSQL(SQL)
                    '
                    ' Housekeep potential issue where a folder deleted and file does not show
                    '
                    SQL = "update cclibraryfiles" _
                            & " Set folderid=null" _
                            & " where id in" _
                            & " (" _
                            & " select c.id from (cclibraryfiles c left join cclibraryfolders p on p.id=c.folderid)" _
                            & " where p.ID Is Null" _
                            & " and c.folderid is not null" _
                            & " and c.folderid<>0" _
                            & " )"
                    Call cp.Db.ExecuteSQL(SQL)
                    '
                    ' ----- Begin output
                    '
                    Dim rnbutton As String = "Button"
                    Dim ButtonExit As String
                    If (SourceMode = SourceModeFromDownloadRequest) Or (SourceMode = SourceModeFromLinkDialog) Then
                        ButtonExit = cp.Html.Button(rnbutton, ButtonClose, , "windowcloseID")
                    Else
                        ButtonExit = cp.Html.Button(rnbutton, ButtonCancel)
                    End If
                    Dim ButtonBar As String = ""
                    If AllowLocalFileAdd Then
                        If currentFolderHasModifyAccess Then
                            ButtonBar = "<div class=ccAdminButtonBar>" _
                                    & ButtonExit _
                                    & cp.Html.Button(rnbutton, ButtonApply) _
                                    & cp.Html.Button(rnbutton, ButtonDelete, RequestNameButton, "returnDeleteCheckID") _
                                    & "</div>"
                        Else
                            ButtonBar = "<div class=ccAdminButtonBar>" _
                                    & cp.Html.Button(rnbutton, ButtonApply) _
                                    & "</div>"
                        End If
                    End If

                    'result = result & genericController.htmlHidden("FolderID", currentFolderID) 
                    result = result & ButtonBar

                    Dim JumpSelect As String = ""
                    JumpSelect = GetJumpFolderPathSelect(cp, currentFolderID, topFolderPath)
                    result = result & "<div style=""padding:10px;"">" & GetParentFoldersLink(cp, topFolderPath, topFolderID, currentFolderID, currentFolderID, cp.Doc.RefreshQueryString, "") & "</div>"
                    If JumpSelect <> "" Then
                        result = result & "<div style=""padding:10px;padding-top:0px"">" & "Jump to&nbsp;" & JumpSelect & "</div>"
                    End If
                    '
                    ' From here down the form divides into FormFolder and FormDetails
                    '
                    Dim FormDetails As String = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr class=""headRow"">"
                    If AllowSelectColumn Then
                        FormDetails = FormDetails & GetForm_HeaderCell(cp, "center", "10", "Select<BR>" & Image10)
                    End If
                    If AllowEditColumn Then
                        FormDetails = FormDetails & GetForm_HeaderCell(cp, "center", "15", "Edit<br>" & Image15)
                    End If
                    If AllowPlaceColumn Then
                        FormDetails = FormDetails & GetForm_HeaderCell(cp, "center", "15", "Place<br>" & Image15)
                    End If
                    FormDetails = FormDetails _
                            & GetForm_HeaderCell(cp, "left", "20", "&nbsp;<BR>" & Image20) _
                            & GetForm_HeaderCell(cp, "left", "20%", "Name<br>" & Image20) _
                            & GetForm_HeaderCell(cp, "left", "50%", "Description<br>" & Image15) _
                            & GetForm_HeaderCell(cp, "center", "50", "Size<br>" & Image50) _
                            & GetForm_HeaderCell(cp, "center", "50", "Modified&nbsp;&nbsp;<br>" & Image50) _
                            & "</tr>"
                    '
                    ' ----- Select the Folder Rows
                    '
                    Dim Criteria As String = "((ParentID is null)or(ParentID=0))"
                    '
                    If currentFolderID <> 0 Then
                        Call cp.Doc.AddRefreshQueryString("FolderID", currentFolderID.ToString())
                    End If
                    '
                    Dim SortField As String = cp.Doc.GetText("sortfield")
                    If SortField = "" Then
                        SortField = "Name"
                    End If
                    Call cp.Doc.AddRefreshQueryString("SortField", SortField)
                    '
                    Dim SortDirection As Integer = cp.Doc.GetInteger("sortdirection")
                    If SortDirection <> 0 Then
                        Call cp.Doc.AddRefreshQueryString("SortDirection", SortDirection.ToString())
                    End If
                    '
                    If SortDirection <> 0 And SortField <> "" Then
                        SortField = SortField & " DESC"
                    End If
                    '
                    Dim parentFolder As LibraryFolderModel = Nothing

                    Dim parentFolderID As Integer
                    If currentFolderID <> 0 Then
                        '
                        ' ----- FolderID given, lookup record and get ParentID
                        '       Note that allowupfolder allows users to "up" past top if they set it manually
                        '       Fix this when security is added
                        '
                        folder = LibraryFolderModel.create(cp, currentFolderID)
                        If (folder IsNot Nothing) Then
                            parentFolderID = folder.ParentID
                        End If
                        parentFolder = LibraryFolderModel.create(cp, parentFolderID)
                        Criteria = "(ParentID=" & KmaEncodeSQLNumber(cp, currentFolderID) & ")"
                    ElseIf topFolderPath <> "" Then
                        '
                        ' ----- Rootfolder given, lookup record and get ParentID
                        '
                        folder = LibraryFolderModel.createByName(cp, topFolderPath)
                        If (folder IsNot Nothing) Then
                            parentFolderID = 0
                            currentFolderID = folder.id
                            Call cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString())
                        End If
                        parentFolder = LibraryFolderModel.create(cp, parentFolderID)
                        Criteria = "(ParentID=" & KmaEncodeSQLNumber(cp, currentFolderID) & ")"
                    Else
                        '
                        ' ----- Use Root as top (no record)
                        parentFolder = LibraryFolderModel.create(cp, parentFolderID)
                        '
                    End If
                    '
                    ' ----- Output the page
                    '
                    RowCount = 0
                    hint = "700"
                    If True Then
                        '
                        ' ----- List out the folders
                        Dim folderList As List(Of LibraryFolderModel) = LibraryFolderModel.createList(cp, Criteria, SortField)
                        Dim IconLink As String
                        Dim EditLink As String
                        Dim ModifiedDate As Date
                        For Each folder In folderList
                            Dim ChildFolderName As String = folder.name
                            If ChildFolderName = "" Then
                                ChildFolderName = "[no name]"
                            End If
                            EditLink = ""
                            If AllowFolderAuthoring Then
                                EditLink = adminUrl(cp) & "?cid=" & FolderCID & "&id=" & folder.id & "&af=4" & "&aa=2&depth=1"
                            End If
                            IconLink = cp.Utils.ModifyQueryString(cp.Doc.RefreshQueryString, "folderid", CStr(folder.id))
                            ModifiedDate = folder.ModifiedDate
                            If ModifiedDate <= Date.MinValue Then
                                ModifiedDate = folder.DateAdded
                            End If
                            Dim ChildFolderID As Integer
                            FormDetails = FormDetails & GetFormRow_ChildFolders(cp, IconFolderClosed, IconLink, "", ChildFolderName, "", ModifiedDate, RowCount, EditLink, folder.Description, "CHILD", "", "", "", "", "", 0, folder.id, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn)
                            RowCount = RowCount + 1
                        Next

                        '
                        ' Lookup the files in the folder
                        '
                        hint = "720"
                        If currentFolderID = 0 Then
                            Criteria = "((FolderID is null)or(FolderID=0))"
                        Else
                            Criteria = "(FolderID=" & KmaEncodeSQLNumber(cp, currentFolderID) & ")"
                        End If
                        'FieldList = "ID,Name,ModifiedDate,Filename,Width,Height,DateAdded,Description,AltText,FileTypeID,FileSize,AltSizeList"
                        If currentFolderID = 0 Then
                            Criteria = "((FolderID is null)or(FolderID=0))"
                        Else
                            Criteria = "(FolderID=" & KmaEncodeSQLNumber(cp, currentFolderID) & ")"
                        End If
                        Dim fileList As List(Of LibraryFileModel) = LibraryFileModel.createList(cp, Criteria, SortField)
                        For Each file In fileList
                            Dim UpdateRecord As Boolean = False
                            Dim ResourceRecordID As Integer = file.id
                            Dim RecordName As String = file.name
                            ModifiedDate = file.ModifiedDate
                            Dim Filename As String = file.Filename
                            Dim ImageWidthText As String = file.Width
                            Dim ImageHeightText As String = file.Height
                            If ModifiedDate <= Date.MinValue Then
                                ModifiedDate = file.DateAdded
                            End If
                            Dim Description As String = file.Description
                            Dim ImageAlt As String = file.AltText
                            FileTypeID = file.FileTypeID
                            Dim fileSize As Integer = file.FileSize
                            AltSizeList = file.AltSizeList
                            '
                            Dim ImageSrc As String = cp.Site.FilePath & Replace(Filename, "\", "/")
                            '
                            Dim DotPosition As Integer = InStrRev(ImageSrc, ".")
                            If DotPosition = 0 Then
                                FileExtension = ""
                                FilenameNoExtension = ""
                            Else
                                FileExtension = UCase(Mid(ImageSrc, DotPosition + 1))
                                FilenameNoExtension = Mid(ImageSrc, 1, DotPosition - 1)
                            End If
                            '
                            If FileTypeID = 0 Then
                                FileTypeID = GetFileTypeID(cp, ImageSrc)
                                If FileTypeID <> 0 Then
                                    UpdateRecord = True
                                End If
                            End If
                            '
                            ' if no name given, use the filename
                            '
                            If RecordName = "" Then
                                If ImageSrc = "" Then
                                    RecordName = "[no name]"
                                Else
                                    DotPosition = InStrRev(ImageSrc, "/")
                                    If DotPosition = 0 Then
                                        RecordName = ImageSrc
                                    Else
                                        RecordName = Mid(ImageSrc, DotPosition + 1)
                                    End If
                                End If

                            End If
                            file.name = RecordName

                            '
                            Dim ResourceHref As String = ""
                            IconLink = ""
                            If AllowFileAuthoring Then
                                EditLink = adminUrl(cp) & "?cid=" & FileCID & "&id=" & ResourceRecordID & "&af=4" & "&aa=2&depth=1"
                            Else
                                EditLink = ""
                            End If
                            Dim ThumbNailSrc As String
                            '
                            ' create thumbnail
                            '
                            If AllowThumbnails Then
                                ThumbNailSrc = ImageSrc
                                If (FilenameNoExtension <> "") And (AltSizeList <> "") Then
                                    Dim AltSizes() As String = Split(AltSizeList, vbCrLf)
                                    Dim BestFitHeight As Integer = 9999999
                                    Dim BestFitAltSize As String = ""
                                    For Ptr = 0 To UBound(AltSizes)
                                        '
                                        ' Find the smallest image larger then height 50
                                        '
                                        Dim AltSize As String = Trim(AltSizes(Ptr))
                                        If AltSize <> "" Then
                                            Pos = InStr(AltSize, "x")
                                            If Pos > 0 Then
                                                Dim AltSizeHeight As Integer = cp.Utils.EncodeInteger(Mid(AltSize, Pos + 1))
                                                If AltSizeHeight >= 50 And AltSizeHeight < BestFitHeight Then
                                                    BestFitHeight = AltSizeHeight
                                                    BestFitAltSize = AltSize
                                                End If
                                            End If
                                        End If
                                    Next
                                    If BestFitAltSize <> "" Then
                                        ThumbNailSrc = FilenameNoExtension & "-" & BestFitAltSize & "." & FileExtension
                                    End If
                                    '
                                    '
                                    '
                                End If
                            End If
                            ' get file size
                            '
                            'FileSize = 0
                            If fileSize = 0 Then
                                Dim Pathname As String = cp.Site.PhysicalFilePath & Replace(Filename, "/", "\")
                                fileSize = GetFileSize(cp, Pathname)
                                If fileSize <> 0 Then
                                    UpdateRecord = True
                                End If
                            End If
                            '
                            '
                            '
                            If UpdateRecord Then
                                Call cp.Db.ExecuteSQL("update cclibraryFiles set FileTypeID=" & FileTypeID & ",filesize=" & fileSize & " where ID=" & ResourceRecordID)
                            End If
                            '
                            ImageSrc = kmaEncodeURL(cp, ImageSrc)
                            Dim IconOnClick As String = ""
                            FormDetails = FormDetails & GetFormRow_Files(cp, fileSize, IconLink, IconOnClick, RecordName, ImageSrc, ModifiedDate, RowCount, EditLink, Description, FileExtension, RecordName, ImageSrc, ImageAlt, ImageWidthText, ImageHeightText, ResourceRecordID, currentFolderID, AllowThumbnails, FileTypeFilter, ThumbNailSrc, SourceMode, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn)
                            RowCount = RowCount + 1
                        Next
                        '
                        ' ----- If nothing found, print no files found
                        If RowCount = 0 Then
                            FormDetails = FormDetails & "<tr class=""listRow""><td class=""center"">" & IconSpacer & "</td><td class=""left"" colspan=" & ColumnCnt - 1 & ">no folders or files were found</td></tr>"
                            RowCount = RowCount + 1
                        End If
                    End If
                    '
                    ' Fill out the table to MinRows
                    '
                    hint = "800"
                    For RowCount = RowCount To iMinRows
                        FormDetails = FormDetails & GetFormRow_Blank(cp, "", "", "", "", "", Nothing, RowCount, "", "", "BLANK", "", "", "", "", "", 0, currentFolderID, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn, ColumnCnt)
                    Next
                    '
                    ' Upload link
                    '
                    If AllowLocalFileAdd Then
                        '
                        ' Upload Form
                        '
                        FormDetails = FormDetails & GetFormRow_Options(cp, currentFolderID, topFolderPath, ColumnCnt, IsContentManagerFiles, IsContentManagerFolders, currentFolderHasModifyAccess)
                        RowCount = RowCount + 1
                    End If
                    '
                    ' Bottom border
                    '
                    FormDetails = FormDetails & "<tr class=""border""><td class=""border"" Colspan=" & (ColumnCnt) & ">" & cp.Html.div("&nbsp;") & "</td></tr>"
                    FormDetails = FormDetails & "</table>"
                    '
                    ' Create the FormFolders
                    '
                    Dim FormFolders As String = GetRLNav(cp, currentFolderID, topFolderPath, topFolderID)
                    FormFolders = "<div class=""rlnav"">" & FormFolders & "</div>"
                    'FormFolders = Main.GetPanelRev(FormFolders)
                    '
                    ' Assemble the form
                    '
                    hint = "900"
                    result = result & "<table border=0 cellpadding=0 cellspacing=0 width=""100%""><tr>"
                    If Not blockFolderNavigation Then
                        result = result & "<td class=""nav ccPanel3DInput"">" & FormFolders & "<BR><img src=/ResourceLibrary/spacer.gif width=140 height=1></td>"
                        result = result & "<td class=""navBorder ccPanel3D""><img src=/ResourceLibrary/spacer.gif width=5 height=1></td>"
                    End If
                    result = result & "<td class=""content"">" & FormDetails & "</td>"
                    result = result & "</tr></Table>"
                    result = result & ButtonBar
                    result = result & htmlHidden("RowCount", RowCount)
                    result = cp.Html.Form(result)
                End If
                '
                result = "<div class=""ccLibrary"">" & result & "</div>"
                '
                ' Help Link
                '
                'result = Main.GetHelpLink(42, "Using the Resource Library", "The Resource Library is a convenient place to store reusable content, such as images and downloads. Objects in the Library can be placed on any page. The Library itself can be added to any page on your site.") & GetForm
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        ' Returns the Resource Library Row HTML.
        '=================================================================================
        '
        Private Function GetFormRow_Folders(cp As CPBaseClass, ignore0 As String, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Integer, EditLink As String, Description As String, FileType As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Integer, FolderID As Integer, AllowEditColumn As Boolean, AllowPlaceColumn As Boolean) As String
            Dim result As String = ""
            '
            Try
                Dim RowClass As String
                '
                If (RowCount Mod 2) = 0 Then
                    RowClass = "ccPanelRowOdd"
                Else
                    RowClass = "ccPanelRowEven"
                End If
                '

                Dim CellStart As String = "<td class=""left ccAdminSmall " & RowClass & """>"
                Dim CellStartRight As String = "<td class=""right ccAdminSmall " & RowClass & """>"
                Dim CellStart2 As String = "<td class=""left ccAdminSmall " & RowClass & """>"
                Dim CellStart5 As String = "<td class=""left ccAdminSmall " & RowClass & """>"
                Dim CellEnd As String = "</td>"
                Dim DateString As String
                '
                If ModifiedDate <= Date.MinValue Then
                    DateString = "&nbsp;"
                Else
                    DateString = FormatDateTime(ModifiedDate, vbShortDate)
                End If
                '
                result = result & "<tr class=""row " & RowClass & """>"
                result = result & CellStart & "&nbsp;" & CellEnd
                If AllowEditColumn Then
                    result = result & CellStart & "&nbsp;" & CellEnd
                End If
                If AllowPlaceColumn Then
                    result = result & CellStart & "&nbsp;" & CellEnd
                    'Else
                    '    result = result & CellStart & "&nbsp;" & CellEnd
                End If
                result = result & CellStart & "<A href=""?" & cp.Utils.EncodeUrl(IconLink) & """>" & IconFolderOpen & "</A>" & CellEnd
                result = result & CellStart & Name & CellEnd
                result = result & CellStart & Description & CellEnd
                result = result & CellStart & "&nbsp;" & CellEnd
                result = result & CellStartRight & DateString & CellEnd
                result = result & "</tr>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        ' Returns the Resource Library Row HTML.
        '=================================================================================
        '
        Private Function GetFormRow_ChildFolders(cp As CPBaseClass, ignore0 As String, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Integer, EditLink As String, Description As String, FileType As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Integer, FolderID As Integer, AllowEditColumn As Boolean, AllowPlaceColumn As Boolean, AllowSelectColumn As Boolean) As String
            Dim result As String = ""
            '
            Try
                '
                Dim RowClass As String
                '
                If (RowCount Mod 2) = 0 Then
                    RowClass = "ccPanelRowOdd"
                Else
                    RowClass = "ccPanelRowEven"
                End If
                '
                Dim CellStart As String = vbCrLf & "<td class=""left ccAdminSmall"">"
                Dim CellStartCenter As String = vbCrLf & "<td class=""center ccAdminSmall"">"
                Dim CellStartRight As String = vbCrLf & "<td class=""right ccAdminSmall"">"
                Dim CellEnd As String = "</td>"
                Dim DateString As String
                '
                If ModifiedDate <= Date.MinValue Then
                    DateString = "&nbsp;"
                Else
                    DateString = FormatDateTime(ModifiedDate, vbShortDate)
                End If
                If Description = "" Then
                    Description = "&nbsp;"
                End If
                '
                GetFormRow_ChildFolders = GetFormRow_ChildFolders & vbCrLf & "<tr class=""listRow"" ID=""Row" & RowCount & """>"
                If AllowSelectColumn Then
                    GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartCenter & "<input type=checkbox ID=Select" & RowCount & " name=Row" & RowCount & " value=1 onClick=""RLRowClick(this.checked,'Row" & RowCount & "');"">" & htmlHidden("Row" & RowCount & "FolderID", FolderID) & CellEnd
                End If
                If AllowEditColumn Then
                    If EditLink <> "" Then
                        GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartCenter & "<A href=""" & EditLink & """>" & IconFolderEdit & "</A>" & CellEnd
                    Else
                        GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStart & "&nbsp;" & CellEnd
                    End If
                End If
                If AllowPlaceColumn Then
                    GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartCenter & IconNoFile & CellEnd
                    'Else
                    '    GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartCenter & IconNoFile & CellEnd
                End If
                GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartCenter & "<A href=""?" & IconLink & """>" & IconFolderClosed & "</A>" & CellEnd
                GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStart & "<A href=""?" & IconLink & """>" & Name & "</A>" & CellEnd
                GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStart & Description & CellEnd
                GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartRight & "&nbsp;" & CellEnd
                GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartRight & DateString & CellEnd
                GetFormRow_ChildFolders = GetFormRow_ChildFolders & "</tr>"
                result = GetFormRow_ChildFolders
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        ' Returns the Resource Library Row HTML.
        '=================================================================================
        '
        Private Function GetFormRow_Files(cp As CPBaseClass, fileSize As Integer, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Integer, EditLink As String, Description As String, FilenameExt As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Integer, FolderID As Integer, AllowThumbnails As Boolean, FileTypeFilter As String, ThumbNailSrc As String, SourceMode As Integer, AllowEditColumn As Boolean, AllowPlaceColumn As Boolean, AllowSelectColumn As Boolean) As String
            Dim result As String = ""
            '
            Try
                Dim RowClass As String
                '
                If (RowCount Mod 2) = 0 Then
                    RowClass = "ccPanelRowOdd"
                Else
                    RowClass = "ccPanelRowEven"
                End If
                '
                Dim CellStart As String = vbCrLf & "<td class=""left ccAdminSmall"">"
                Dim CellStartCenter As String = vbCrLf & "<td class=""center ccAdminSmall"">"
                Dim CellStartRight As String = vbCrLf & "<td class=""right ccAdminSmall"">"
                Dim CellEnd As String = "</td>"
                Dim DateString As String
                '
                If ModifiedDate <= Date.MinValue Then
                    DateString = "&nbsp;"
                Else
                    DateString = FormatDateTime(ModifiedDate, vbShortDate)
                End If
                '
                ' Determine Icons and actions
                '
                Dim AllowPlace As Boolean
                AllowPlace = False
                Dim IconIMG As String
                Dim IconFilename As String
                Dim IsImage As Boolean
                Dim IsVideo As Boolean
                Dim IsFlash As Boolean
                Dim IsDownload As Boolean
                Dim Downloadfilename As String
                Dim FileTypeName As String
                Dim TestFileTYpe As String
                Dim FileTypeFound As Boolean
                Dim MediaIMG As String = ""
                If iconFiles.Count <= 0 Then
                    IconIMG = IconImage
                Else
                    TestFileTYpe = "," & UCase(Replace(FilenameExt, ".", "")) & ","
                    For Each iconFile As FileType In iconFiles
                        Dim FileTypePtr As Integer
                        If InStr(1, "," & iconFiles(FileTypePtr).ExtensionList & ",", UCase(TestFileTYpe), vbTextCompare) <> 0 Then
                            With iconFile
                                FileTypeName = .Name
                                IsImage = .IsImage
                                IsVideo = .IsVideo
                                IsFlash = .IsFlash
                                Dim IsMedia As Boolean = IsImage Or IsVideo Or IsFlash
                                '
                                ' 4/15/08 - if no filter, show everything
                                '
                                'MediaIMG = IconNoFile

                                '                        If FileTypeFilter = "image" And IsImage Then
                                '                            MediaIMG = "<img src=""/ResourceLibrary/IconImagePlace2322.gif"" width=23 height=22 border=0 alt=""Place this image on the page"">"
                                '                            AllowPlace = True
                                '                        ElseIf FileTypeFilter = "media" And IsVideo Then
                                '                            MediaIMG = "<img src=""/ResourceLibrary/IconImagePlace2322.gif"" width=23 height=22 border=0 alt=""Place this video on the page"">"
                                '                            AllowPlace = True
                                '                        ElseIf FileTypeFilter = "flash" And IsVideo Then
                                '                            MediaIMG = "<img src=""/ResourceLibrary/IconImagePlace2322.gif"" width=23 height=22 border=0 alt=""Place this flash on the page"">"
                                '                            AllowPlace = True
                                '                        Else
                                '                            MediaIMG = IconNoFile
                                '                            AllowPlace = False
                                '                        End If
                                If .MediaIconFilename <> "" Then
                                    MediaIMG = "<img src=""" & .MediaIconFilename & """ width=23 height=22 border=0 alt=""Place this flash on the page"">"
                                End If
                                IsDownload = .IsDownload
                                Downloadfilename = .DownloadIconFilename
                                IconFilename = .IconFilename
                                If IconFilename = "" Then
                                    IconFilename = "/ResourceLibrary/IconDefault.gif"
                                End If
                                IconIMG = "<img src=""" & IconFilename & """ border=""0"" width=""22"" height=""23"" alt=""" & .Name & """>"
                            End With
                            FileTypeFound = True
                            Exit For
                        End If
                    Next
                    'For FileTypePtr = 0 To IconFileCnt - 1

                    'Next
                End If
                '
                If Not FileTypeFound Then
                    If FilenameExt = "PNG" Then
                        IsImage = True
                    ElseIf FilenameExt = "JPG" Then
                        IsImage = True
                    ElseIf FilenameExt = "GIF" Then
                        IsImage = True
                    Else
                        IsImage = False
                    End If
                    FileTypeName = TestFileTYpe
                    IsVideo = False
                    IsFlash = False
                    Dim Mediafilename As String = ""
                    IsDownload = True
                    Downloadfilename = "/ResourceLibrary/IconDefaultDownload.gif"
                    IconFilename = "/ResourceLibrary/IconFile.gif"
                    IconIMG = IconOther
                    MediaIMG = IconNoFile
                End If
                AllowPlace = False
                If FileTypeFilter = "image" Then
                    If IsImage Then
                        AllowPlace = True
                    End If
                ElseIf FileTypeFilter = "media" Then
                    If IsVideo Then
                        AllowPlace = True
                    End If
                ElseIf FileTypeFilter = "flash" Then
                    If IsFlash Then
                        AllowPlace = True
                    End If
                Else
                    '
                    ' no filter - place anything
                    '
                    AllowPlace = True
                End If
                If AllowPlace And MediaIMG = "" Then
                    MediaIMG = "<img src=""/ResourceLibrary/IconImagePlace2322.gif"" width=23 height=22 border=0 alt=""Place this file on the page"">"
                End If
                '
                '   Output the row
                '
                GetFormRow_Files = GetFormRow_Files & vbCrLf & "<tr class=""listRow"" ID=""Row" & RowCount & """>"
                If AllowSelectColumn Then
                    GetFormRow_Files = GetFormRow_Files & CellStartCenter & "<input type=checkbox ID=Select" & RowCount & " name=Row" & RowCount & " value=1 onClick=""RLRowClick(this.checked,'Row" & RowCount & "');"">" & htmlHidden("Row" & RowCount & "FileID", RecordID) & CellEnd
                End If
                '
                ' ----- Edit Column
                '
                If AllowEditColumn Then
                    If EditLink <> "" Then
                        GetFormRow_Files = GetFormRow_Files & CellStartCenter & "<A href=""" & EditLink & """>" & IconFileEdit & "</A>" & CellEnd
                    Else
                        GetFormRow_Files = GetFormRow_Files & CellStart & "&nbsp;" & CellEnd
                    End If
                End If
                '
                ' ----- Place Column
                '
                If Not AllowPlaceColumn Then
                    '
                    ' hide column
                    '
                ElseIf (Not AllowPlace) Then
                    '
                    ' Can not select resources - display dot
                    '
                    GetFormRow_Files = GetFormRow_Files & CellStartCenter & IconNoFile & CellEnd
                Else
                    '
                    Dim ImageLink As String
                    Dim JSCopy As String
                    '
                    ' Allow selection of resources to be placed on the opening pages
                    '
                    If SelectLinkObjectName <> "" Then
                        '
                        ' return the objects URL to the input element with ID=SelectLinkObjectName
                        '
                        JSCopy = kmaEncodeJavascript(cp, ResourceLink)
                        ImageLink = "<img src=""/ResourceLibrary/resourceLink1616.gif"" border=""0"" width=""16"" height=""16"" alt=""Place a link to this resource"" title=""Place a link to this resource"" valign=""absmiddle"">"
                        GetFormRow_Files = GetFormRow_Files & CellStartCenter & "<a href=""#"" onClick=""var e=window.opener.document.getElementById('" & SelectLinkObjectName & "');e.value='" & JSCopy & "'; window.close();"">" & ImageLink & "</A>" & CellEnd
                    ElseIf SourceMode = SourceModeFromDownloadRequest Then
                        '
                        ' return a simple download
                        '
                        If IsDownload Then
                            JSCopy = Downloadfilename
                            JSCopy = Replace(JSCopy, "\", "\\")
                            JSCopy = kmaEncodeJavascript(cp, JSCopy)
                            ImageLink = "<img src=""/ResourceLibrary/IconDownload2.gif"" border=""0"" width=""23"" height=""22"" alt=""Link to this resource"" title=""Link to this resource"" valign=""absmiddle"">"
                            GetFormRow_Files = GetFormRow_Files & CellStartCenter & "<a href=""#"" onClick=""window.opener.InsertDownload( '" & RecordID & "','" & SelectResourceEditorObjectName & "','" & JSCopy & "'); window.close();"">" & ImageLink & "</A>" & CellEnd
                        Else
                            GetFormRow_Files = GetFormRow_Files & CellStartCenter & IconNoFile & CellEnd
                        End If
                    ElseIf SourceMode = SourceModeFromLinkDialog Then
                        '
                        ' Return the file as a url to the editor dialog
                        '
                        If AllowPlace Then
                            JSCopy = kmaEncodeJavascript(cp, ResourceLink)
                            Dim JSClose As String = "" _
                            & " if(navigator.appName.indexOf('Microsoft')!=-1) {window.returnValue='" & JSCopy & "'}" _
                            & " else{window.opener.setAssetValue('" & JSCopy & "')}" _
                            & " self.close();"
                            GetFormRow_Files = GetFormRow_Files & CellStartCenter & "<a href=""#"" onClick=""" & JSClose & """ >" & MediaIMG & "</A>" & CellEnd
                        Else
                            GetFormRow_Files = GetFormRow_Files & CellStartCenter & IconNoFile & CellEnd
                        End If
                    End If
                End If
                NameLink = cp.Utils.DecodeUrl(NameLink)

                GetFormRow_Files = GetFormRow_Files & CellStartCenter & IconIMG & CellEnd
                GetFormRow_Files = GetFormRow_Files & CellStart & "<a href=""" & NameLink & """ target=""_blank"">" & Name & "</A>" & CellEnd
                '
                If Description = "" Then
                    Description = "&nbsp;"
                End If
                If AllowThumbnails And IsImage Then
                    'If AllowThumbnails And (UCase(FileTypeName) = "IMAGE") Then
                    GetFormRow_Files = GetFormRow_Files _
            & CellStart _
            & "<a href=""" & NameLink & """ target=""_blank"">" _
            & "<img src=""" & ThumbNailSrc & """ height=""50""  vspace=""0"" hspace=""10"" style=""vertical-align:middle;border:0;"">" _
            & "</a>" _
            & Description _
            & CellEnd
                Else
                    GetFormRow_Files = GetFormRow_Files _
            & CellStart _
            & Description _
            & CellEnd
                End If
                '
                If fileSize > 10000 Then
                    GetFormRow_Files = GetFormRow_Files & CellStartRight & Int(fileSize / 1024) & "&nbsp;KB&nbsp;" & CellEnd
                Else
                    GetFormRow_Files = GetFormRow_Files & CellStartRight & fileSize & "&nbsp;" & CellEnd
                End If
                '
                GetFormRow_Files = GetFormRow_Files & CellStartRight & DateString & CellEnd
                GetFormRow_Files = GetFormRow_Files & "</tr>"
                result = GetFormRow_Files
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function


        '
        '=================================================================================
        ' Returns the Resource Library Row HTML.
        '=================================================================================
        '
        Private Function GetFormRow_Blank(cp As CPBaseClass, ignore0 As String, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Integer, EditLink As String, Description As String, FileType As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Integer, FolderID As Integer, AllowEditColumn As Boolean, AllowPlaceColumn As Boolean, AllowSelectColumn As Boolean, ColumnCnt As Integer) As String

            '
            GetFormRow_Blank = vbCrLf & vbTab & "<tr class=""listRow""><td class=""left""><img height=""23"" width=""22"" src=""/wwwroot/ResourceLibrary/spacer.gif""></td><td class=""left"" colspan=""" & ColumnCnt - 1 & """>&nbsp;</td></tr>"
            '
        End Function
        '
        '=================================================================================
        ' Returns the Resource Library Row HTML.
        '=================================================================================
        '
        Private Function GetFormRow_Options(cp As CPBaseClass, FolderID As Integer, topFolderPath As String, ColumnCnt As Integer, IsContentManagerFiles As Boolean, IsContentManagerFolders As Boolean, hasModifyAccess As Boolean) As String
            Dim result As String = ""
            Try
                Dim FolderCell As String = ""
                '
                ' Inner Cell
                '
                If hasModifyAccess Then
                    '
                    ' if you have viewaccess to the folder, you can see it
                    ' if you have modifyaccess to the folder, you can upload to it and create subfolders in it
                    '
                    'If IsContentManagerFolders Then
                    FolderCell = "" _
                    & "<table id=""AddFolderTable"" border=""0"" cellpadding=""10"" cellspacing=""1"" width=""100%"">" _
                    & "<tr>"
                    FolderCell = FolderCell _
                    & "<td class=""left"" align=""left"" colspan=2>" & kmaAddSpan("Add Folder&nbsp;", "ccAdminSmall") & "<BR><img src=""/ResourceLibrary/spacer.gif"" width=""230"" height=""1""></td>" _
                    & "<td class=""left"" Width=""99%"" align=""left"">" & kmaAddSpan("Description&nbsp;", "ccAdminSmall") & "<BR><img src=""/ResourceLibrary/spacer.gif"" width=""100"" height=""1""></td>" _
                    & "</tr><tr>" _
                    & "<td class=""left"" Width=""30"" align=""right"">1&nbsp;<img src=/ResourceLibrary/spacer.gif width=30 height=1></td>" _
                    & "<td class=""left"" align=""left""><INPUT TYPE=""Text"" NAME=""FolderName.1"" SIZE=""30""></td>" _
                    & "<td class=""left"" align=""left""><INPUT TYPE=""Text"" NAME=""FolderDescription.1"" SIZE=""40""></td>" _
                    & "</tr>"
                    FolderCell = FolderCell _
                    & "</Table>" _
                    & "<table border=""0"" cellpadding=""10"" cellspacing=""1"" width=""100%"">" _
                    & "<tr><td class=""left"" Width=""30""><img src=/ResourceLibrary/spacer.gif width=30 height=1></td><td align=""left""><a href=""#"" onClick=""InsertFolderRow(); return false;"">+ Add more folders</a></td></tr>" _
                    & "</Table>" & htmlHidden("AddFolderCount", 1, "", "AddFolderCount")
                End If
                Dim FileCell As String = ""
                If hasModifyAccess Then
                    FileCell = FileCell _
                    & "<table id=""UploadInsert"" border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%"">" _
                    & "<tr>"
                    FileCell = FileCell _
                    & "<td class=""left"" align=""left"" colspan=2>" & kmaAddSpan("Add Files&nbsp;", "ccAdminSmall") & "<BR><img src=""/ResourceLibrary/spacer.gif"" width=""230"" height=""1""></td>" _
                    & "<td class=""left"" Width=""100"" align=""left"">" & kmaAddSpan("Name&nbsp;", "ccAdminSmall") & "<BR><img src=""/ResourceLibrary/spacer.gif"" width=""100"" height=""1""></td>" _
                    & "<td class=""left"" Width=""100"" align=""left"">" & kmaAddSpan("Description&nbsp;", "ccAdminSmall") & "<BR><img src=""/ResourceLibrary/spacer.gif"" width=""100"" height=""1""></td>" _
                    & "<td class=""left"" Width=""99%"">&nbsp;</td>" _
                    & "</tr><tr>" _
                    & "<td class=""left"" Width=""30"" align=""right"">1&nbsp;<BR><img src=/ResourceLibrary/spacer.gif width=30 height=1></td>" _
                    & "<td class=""left"" Width=""200"" align=""right""><INPUT TYPE=""file"" name=""LibraryUpload.1""><BR><img src=/ResourceLibrary/spacer.gif width=200 height=1></td>" _
                    & "<td class=""right"" align=""right""><INPUT TYPE=""Text"" NAME=""LibraryName.1"" SIZE=""25""></td>" _
                    & "<td class=""right"" align=""right""><INPUT TYPE=""Text"" NAME=""LibraryDescription.1"" SIZE=""39""></td>" _
                    & "<td class=""left"">&nbsp;</td>" _
                    & "</tr>"
                    FileCell = FileCell _
                    & "</Table>" _
                    & "<table border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%"">" _
                    & "<tr><td class=""left"" Width=""30""><img src=/ResourceLibrary/spacer.gif width=30 height=1></td><td class=""left"" align=""left""><a href=""#"" onClick=""InsertUploadRow(); return false;"">+ Add more files</a></td></tr>" _
                    & "</Table>" & htmlHidden("LibraryUploadCount", 1, "", "LibraryUploadCount")
                End If
                '
                '
                '
                GetFormRow_Options = "" _
                    & "<div  style=""margin-left:10px;""><img src=""/ResourceLibrary/spacer.gif"" width=""1"" height=""5"">" _
                    & "<BR>" & cp.Html.CheckBox("AllowThumbnails", cp.User.GetBoolean("LibraryAllowthumbnails", "0")) & "&nbsp;Display Thumbnails"
                If cp.User.IsAdmin Or hasModifyAccess Then
                    '
                    Dim moveSelect As String = GetMoveFolderPathSelect(cp, FolderID, topFolderPath)
                    If moveSelect <> "" Then
                        GetFormRow_Options = GetFormRow_Options & "<BR>" & cp.Html.CheckBox("Move", False) & "&nbsp;Move selected files to " & moveSelect
                    End If
                    If FolderCell <> "" Then
                        GetFormRow_Options = GetFormRow_Options & "<BR><BR>" & cp.Html.div(FolderCell)
                    End If
                    If FileCell <> "" Then
                        GetFormRow_Options = GetFormRow_Options & "<BR>" & cp.Html.div(FileCell)
                    End If
                End If
                If GetFormRow_Options <> "" Then
                    GetFormRow_Options = cp.Html.div(GetFormRow_Options)
                    GetFormRow_Options = "<tr><td class=""left"" colspan=" & (ColumnCnt) & ">" & GetFormRow_Options & "</td></tr>"
                End If
                '
                result = GetFormRow_Options
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '
        '
        Private Function GetForm_HeaderCell(cp As CPBaseClass, Align As String, Width As String, Copy As String) As String
            Dim Style As String = "" _
                    & "padding: 3px;" _
                    & "font-size:10px;"
            Dim result As String = "<td WIDTH=""" & Width & """ ALIGN=""" & Align & """ class=ccAdminListCaption style=""" & Style & """>" _
                    & Copy _
                    & "</td>"
            Return result
        End Function
        '
        '
        '
        Private Function IsInFolder(cp As CPBaseClass, topFolderID As Integer, FolderID As Integer, Optional ParentPath As String = "") As Boolean
            Dim result As String = ""
            Try
                '
                Dim cS As Integer
                Dim ParentID As Integer
                '
                If FolderID = 0 Then
                    IsInFolder = False
                ElseIf topFolderID = 0 Then
                    IsInFolder = True
                ElseIf (InStr(1, "," & ParentPath & ",", "," & CStr(FolderID) & ",") <> 0) Then
                    IsInFolder = False
                Else
                    ParentPath = ParentPath & "," & CStr(FolderID)
                    Dim folder As LibraryFolderModel = LibraryFolderModel.create(cp, FolderID)
                    If (folder IsNot Nothing) Then
                        ParentID = folder.ParentID
                    End If
                    If ParentID = 0 Then
                        IsInFolder = False
                    ElseIf ParentID = topFolderID Then
                        IsInFolder = True
                    Else
                        IsInFolder = IsInFolder(cp, topFolderID, ParentID, ParentPath)
                    End If
                End If
                '
                result = IsInFolder
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '
        '
        Private Function GetParentFoldersLink(cp As CPBaseClass, topFolderPath As String, topFolderID As Integer, currentFolderID As Integer, FolderID As Integer, RefreshQS As String, ChildIDList As String) As String
            Dim folderName As String = ""
            Dim result As String = ""
            Try
                '
                If FolderID = 0 Or (FolderID = topFolderID) Then
                    '
                    ' Root folder
                    '
                    folderName = topFolderPath
                    If folderName = "" Then
                        folderName = "Root"
                    End If
                    If currentFolderID = FolderID Then
                        GetParentFoldersLink = "Folder <B>" & folderName & "</B>"
                    Else
                        GetParentFoldersLink = "Folder <a href=?" & RefreshQS & "&FolderID=0>" & folderName & "</a>"
                    End If
                Else
                    Dim LibraryFolder As LibraryFolderModel = LibraryFolderModel.create(cp, "ID=" & FolderID)

                    '

                    Dim ParentID As Integer
                    Dim RecordFound As Boolean
                    'cS = main.OpenCSContent("Library Folders", "ID=" & FolderID, , , , , "Name,ParentID")
                    If Not (LibraryFolder Is Nothing) Then
                        RecordFound = True
                        ParentID = LibraryFolder.ParentID
                        folderName = LibraryFolder.name
                    End If

                    Dim FolderLink As String
                    '
                    If currentFolderID = FolderID Then
                        FolderLink = "<B>" & folderName & "</B>"
                    Else
                        FolderLink = "<a href=?" & RefreshQS & "&FolderID=" & FolderID & ">" & folderName & "</a>"
                    End If
                    '
                    If (Not RecordFound) Or (FolderID = topFolderID) Then
                        '
                        ' call this the top of the tree
                        '
                        If folderName = "" Then
                            folderName = "Root"
                        End If
                        GetParentFoldersLink = "Folder " & FolderLink
                    ElseIf InStr(1, ChildIDList & ",", "," & FolderID & ",") <> 0 Then
                        '
                        ' circular reference - end it here
                        '
                        GetParentFoldersLink = "Folder (Circular Reference) > " & FolderLink
                    ElseIf currentFolderID = ParentID Then
                        '
                        ' circular reference - end it here
                        '
                        GetParentFoldersLink = "Folder " & FolderLink
                    Else
                        GetParentFoldersLink = GetParentFoldersLink(cp, topFolderPath, topFolderID, currentFolderID, ParentID, RefreshQS, ChildIDList & "," & FolderID) & "\" & FolderLink
                        'GetParentFoldersLink = GetParentFoldersLink(cp,topFolderPath, topFolderID, CurrentFolderID, ParentID, RefreshQS, ChildIDList & "," & FolderID) & " > " & FolderLink
                    End If
                End If
                result = GetParentFoldersLink
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result

            Exit Function
        End Function
        '
        '----------------------------------------------------------------------------------------
        '   Get a select menu of all folders with which you have ModifyAccess
        '----------------------------------------------------------------------------------------
        '
        Private Function GetFolderPathSelect(cp As CPBaseClass, topFolderPathID As Integer, topFolderPath As String, RequireModifyAccess As Boolean) As String
            Dim result As String = ""
            Try
                Dim pathRemoveString As String = ""
                Dim pathCaption As String
                '
                'GetFolderPathSelect = FolderSelect
                If GetFolderPathSelect = "" Then
                    '
                    '
                    ' create full paths, set .hasViewAccess
                    '
                    Dim optionCnt As Integer = 0
                    If topFolderPath <> "" Then
                        pathRemoveString = "root\"
                        Dim Pos As Integer = InStrRev(topFolderPath, "\")
                        If Pos > 0 Then
                            pathRemoveString = pathRemoveString & Mid(topFolderPath, 1, Pos - 1)
                        End If
                    End If
                    '
                    ' create select
                    '
                    optionCnt = 0
                    If topFolderPath = "" Then
                        '
                        ' if root folder is top folder, everyone has view access
                        '
                        optionCnt = optionCnt + 1
                        If topFolderPathID = 0 Then
                            '
                            ' if root is current folder, mark it selected
                            '
                            GetFolderPathSelect = GetFolderPathSelect & "<option value=0 selected>Root</option>"
                        Else
                            GetFolderPathSelect = GetFolderPathSelect & "<option value=0>Root</option>"
                        End If
                    End If
                    Dim Ptr As Integer = FolderPathIndex.GetFirstPointer
                    Do While (Ptr >= 0)
                        If folders(Ptr).hasViewAccess And ((Not RequireModifyAccess) Or folders(Ptr).hasModifyAccess) Then
                            Dim PtrFolderID As Integer = folders(Ptr).FolderID
                            pathCaption = Replace(folders(Ptr).FullPath, pathRemoveString, "", , , vbTextCompare)

                            If PtrFolderID = topFolderPathID Then
                                GetFolderPathSelect = GetFolderPathSelect & "<option value=" & PtrFolderID & " selected>" & pathCaption & "</option>"
                            Else
                                GetFolderPathSelect = GetFolderPathSelect & "<option value=" & PtrFolderID & ">" & pathCaption & "</option>"
                            End If
                            optionCnt = optionCnt + 1
                        End If
                        Ptr = FolderPathIndex.GetNextPointer
                    Loop
                    '
                    ' Create Select
                    '
                    If optionCnt <= 1 Then
                        '
                        ' If only one folder, (the current one), return nothing
                        '
                        GetFolderPathSelect = ""
                    Else
                        'If GetFolderPathSelect <> "" Then
                        GetFolderPathSelect = "<select name=FieldName size=1 onChange>" & GetFolderPathSelect & "</select>"
                    End If
                    FolderSelect = GetFolderPathSelect
                End If
                '
                result = FolderSelect
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
            Exit Function
        End Function
        '
        '
        '
        Private Function GetFolderPath(cp As CPBaseClass, targetPtr As Integer, ChildIDList As String) As String
            Dim result As String = ""
            Try
                '
                Dim ParentPtr As Integer
                Dim ParentID As Integer
                Dim FolderID As Integer

                '
                result = folders(targetPtr).Name
                ParentID = folders(targetPtr).parentFolderID
                FolderID = folders(targetPtr).FolderID
                If ParentID = 0 Then
                    '
                    ' At the Root page
                    '
                    result = "Root\" & result
                ElseIf (FolderID = ParentID) Or (InStr(1, "," & ChildIDList & ",", "," & ParentID & ",") <> 0) Then
                    '
                    ' circular reference - Make this a root page b
                    '
                Else
                    For ParentPtr = 0 To UBound(folders)
                        '
                        'todo Folder(parentPtr) throws a null ref this needs to be resolved
                        If (folders(ParentPtr) Is Nothing) Then
                            cp.Utils.AppendLogFile("getfolderPath=6b ******** parentPtr [" & ParentPtr & "]")
                        Else
                            If folders(ParentPtr).FolderID = ParentID Then
                                result = GetFolderPath(cp, ParentPtr, ChildIDList & "," & ParentID) & "\" & result
                                'GetFolderPath = GetFolderPath(ParentPtr, ChildIDList & "," & ParentID) & " > " & GetFolderPath
                                Exit For
                            End If
                        End If
                    Next
                End If
                '

            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
            Exit Function
        End Function
        '
        '
        '
        Private Function GetJumpFolderPathSelect(cp As CPBaseClass, FolderID As Integer, topFolderPath As String) As String
            Dim result As String = ""
            Try
                '
                result = GetFolderPathSelect(cp, FolderID, topFolderPath, False)
                If result <> "" Then
                    result = Replace(result, "FieldName", "JumpFolderID")
                    result = Replace(result, "onChange", "onChange=""QJump(this);"" ")
                    result = Replace(result, "value=", "value=?" & cp.Doc.RefreshQueryString & "&FolderID=")
                    result = "<script language=JavaScript1.2>function QJump(e){var l=e.value;if(l!=''){window.name='RL';window.location.assign(l);}}</script>" & result
                End If
                '
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result

        End Function

        '
        '
        '
        Private Function GetMoveFolderPathSelect(cp As CPBaseClass, FolderID As Integer, topFolderPath As String) As String
            Try
                '
                GetMoveFolderPathSelect = GetFolderPathSelect(cp, FolderID, topFolderPath, True)
                GetMoveFolderPathSelect = Replace(GetMoveFolderPathSelect, "FieldName", "MoveFolderID")
                GetMoveFolderPathSelect = Replace(GetMoveFolderPathSelect, "onChange", "onChange=""var e=getElementById('Move');if(e){e.checked=true};"" ")
                '
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
        End Function
        '
        '=============================================================
        '
        '=============================================================
        '
        Private Function GetRLNav(cp As CPBaseClass, currentFolderID As Integer, topFolderPath As String, topFolderID As Integer) As String
            Dim IsAuthoring As Boolean
            Try
                '
                IsAuthoring = False
                Dim BakeName As String = "RLNav"
                If Not IsAuthoring Then
                    '        GetRLNav = Main.ReadBake(BakeName)
                End If
                If GetRLNav = "" Then
                    Dim LinkBase As String = cp.Doc.RefreshQueryString
                    LinkBase = cp.Utils.ModifyQueryString(LinkBase, "FolderID", "0")

                    '
                    '

                    Dim Tree As New menuTreeClass(cp)
                    If topFolderID = 0 Then
                        Call Tree.AddEntry(CStr(0), CStr(-1), "", "", "?" & LinkBase, "Root")
                    End If
                    If folderCnt > 0 Then
                        Dim Ptr As Integer
                        For Ptr = 0 To folderCnt - 1
                            Dim Id As Integer = folders(Ptr).FolderID
                            If folders(Ptr).hasViewAccess Then
                                'If hasModifyAccessByFolder(Id, topFolderPath) Then
                                Dim ParentID As Integer = folders(Ptr).parentFolderID
                                Dim Caption As String = Replace(folders(Ptr).Name, " ", "&nbsp;")
                                Dim Link As String = "?" & cp.Utils.ModifyQueryString(LinkBase, "FolderID", CStr(Id))
                                Call Tree.AddEntry(CStr(Id), CStr(ParentID), "", "", Link, Caption)
                            End If
                        Next
                    End If
                    GetRLNav = Tree.GetTree(CStr(topFolderID), CStr(currentFolderID))
                    ' Call cp.Response.(BakeName, GetRLNav, "Library Folders")
                End If
                '    '
                '    ' Get topFolderPath
                '    '
                '    If topFolderPath = "" Then
                '        topFolderPath = "Root"
                '    Else
                '        topFolderPath = topFolderPath
                '    End If
                '
                ' open the current node
                '

                'Call main.AddOnLoadJavascript("convertTrees(); expandToItem('tree0','" & currentFolderID & "');")
                Call cp.Doc.AddOnLoadJavascript("convertTrees(); expandToItem('tree0','" & currentFolderID & "');")
                'Link = "?" & LinkBase
                'Link = "<div style=""position:relative;left:-10;margin-bottom:3px;""><a href=" & Link & " style=""text-decoration:none ! important;"">" & topFolderPath & "</a></div>"
                'GetRLNav = Replace(GetRLNav, "<LI ", Link & "<LI ", 1, 1, vbTextCompare)
                ''If CurrentFolderID <> 0 Then
                'GetRLNav = GetRLNav & "<script type=""text/javascript"">convertTrees(); expandToItem('tree0','" & CurrentFolderID & "');</script>"
                ''End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
        End Function
        '
        '
        '
        Private Function AllowFolderAccess(cp As CPBaseClass, FolderID As Integer, ParentID As Integer) As Boolean
            '
            Try
                Dim GrandParentID As Integer
                Dim cs1 As CPCSBaseClass = cp.CSNew()
                '
                If FolderID = 0 Or cp.User.IsAdmin Then
                    AllowFolderAccess = True
                Else
                    Dim LibraryFolderModelList As List(Of Models.LibraryFolderModel) = Models.LibraryFolderModel.AllowFolderAccess(cp, FolderID, ParentID)
                    '
                    If Not AllowFolderAccess And (ParentID <> 0) Then
                        Dim LibraryFolder As LibraryFolderModel = LibraryFolderModel.create(cp, ParentID)
                        If Not (LibraryFolder Is Nothing) Then
                            GrandParentID = LibraryFolder.ParentID
                        End If
                        AllowFolderAccess = AllowFolderAccess(cp, ParentID, GrandParentID)
                    End If
                End If
                '
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
        End Function
        '
        '
        '
        Private Function hasModifyAccessByFolder(cp As CPBaseClass, FolderID As Integer, topFolderPath As String) As Boolean
            Try
                '
                Dim Ptr As Integer
                '
                If FolderID = 86 Then
                    FolderID = FolderID
                End If

                '
                If cp.User.IsAdmin Then
                    '
                    '
                    '
                    hasModifyAccessByFolder = True
                Else
                    '
                    ' Need to check permissions
                    '
                    Call LoadFolders_returnTopFolderId(cp, topFolderPath)
                    If FolderID = 0 Then
                        hasModifyAccessByFolder = True
                    Else
                        Ptr = FolderIdIndex.getPtr(CStr(FolderID))
                        If Ptr >= 0 Then
                            hasModifyAccessByFolder = folders(Ptr).hasModifyAccess
                        End If
                    End If
                End If
                '
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
        End Function
        '
        '
        '
        Private Function LoadFolders_returnTopFolderId(cp As CPBaseClass, topFolderPath As String) As Integer
            Dim topFolderID As Integer
            'Dim cs1 As CPCSBaseClass = cp.CSNew()
            '
            Dim result As String = ""
            Try
                FolderIdIndex = New FastIndexClass
                FolderNameIndex = New FastIndexClass
                FolderPathIndex = New FastIndexClass
                If (True) Then
                    Dim IsAdmin As Boolean = cp.User.IsAdmin
                    Dim lcasetopFolderPath As String = LCase(topFolderPath)
                    '
                    ' Load the folders storage
                    '
                    Dim foldersList As List(Of Models.LibraryFolderModel) = Models.LibraryFolderModel.createList(cp, "")
                    folderCnt = 0
                    If (foldersList.Count > 0) Then
                        '
                        ' Store folders and setup folder index
                        '
                        folderCnt = foldersList.Count
                        ReDim folders(foldersList.Count - 1)
                        Dim Ptr As Integer = 0

                        For Each folder In foldersList
                            folders(Ptr) = New FolderType()
                            If True Then
                                If True Then
                                    Call FolderIdIndex.SetPointer(CStr(folder.id), Ptr)
                                    Call FolderNameIndex.SetPointer(folder.name, Ptr)
                                    With folders(Ptr)
                                        .FolderID = folder.id
                                        .parentFolderID = folder.ParentID
                                        .Name = folder.name
                                        .hasModifyAccess = True
                                        .modifyAccessIsValid = True
                                        .hasViewAccess = True
                                    End With
                                    '
                                    ' FullPath, propigate modifyAccess from parent to folder , ViewAccess
                                    '
                                    With folders(Ptr)
                                        '
                                        ' determine modify access
                                        '
                                        If (Not .modifyAccessIsValid) Then
                                            .hasModifyAccess = LoadFolders_GetModifyAccess(cp, .parentFolderID)
                                            .modifyAccessIsValid = True
                                        End If
                                        '
                                        Dim testFullPath As String = GetFolderPath(cp, Ptr, "")
                                        folders(Ptr).FullPath = testFullPath
                                        'End If
                                        Call FolderPathIndex.SetPointer(testFullPath, Ptr)
                                        '
                                        If InStr(1, testFullPath, "root\" & topFolderPath, vbTextCompare) = 1 Then
                                            '
                                            .hasViewAccess = True
                                        End If
                                    End With
                                    '
                                    '
                                    topFolderID = 0
                                    If topFolderPath <> "" Then
                                        Dim targetFolders() As String = Split(topFolderPath, "\")
                                        Dim targetFolderCnt As Integer = UBound(targetFolders) + 1
                                        topFolderID = loadFolders_getFolderID(cp, targetFolders, targetFolderCnt - 1)
                                        '
                                        ' if topFolderId not found, create the new folder(s) necessary to targetFolderPath
                                        '
                                        Dim targetFolderName As String = ""
                                        If topFolderID = 0 Then
                                            Dim targetFolderId As Integer = 0
                                            For Ptr = 0 To targetFolderCnt - 1
                                                targetFolderName = targetFolders(Ptr)
                                                Dim targetParentFolderID As Integer = targetFolderId
                                                '
                                                ' find or create the folder with this name and this targetParentFolderID
                                                '
                                                Dim testFolderPtr As Integer = FolderNameIndex.getPtr(targetFolders(Ptr))
                                                Do While testFolderPtr >= 0
                                                    Dim testParentID As Integer = folders(testFolderPtr).parentFolderID
                                                    If targetParentFolderID <> testParentID Then
                                                        '
                                                        ' right name but wrong parent, try next
                                                        '
                                                    Else
                                                        '
                                                        ' good match, this as the parent and find the next
                                                        '
                                                        Exit Do
                                                    End If
                                                    testFolderPtr = FolderNameIndex.GetNextPointerMatch(targetFolderName)
                                                Loop
                                                If testFolderPtr >= 0 Then
                                                    targetFolderId = folders(testFolderPtr).FolderID
                                                Else
                                                    '
                                                    ' folder not found, create it with the parent
                                                    '
                                                    'cS = main.InsertCSRecord("Library Folders")

                                                    'If main.IsCSOK(cS) Then
                                                    If Not (folder Is Nothing) Then
                                                        targetFolderId = folder.id
                                                        folder.name = targetFolderName
                                                        folder.ParentID = targetParentFolderID
                                                        'Call main.SetCS(cS, "name", targetFolderName)
                                                        'Call main.SetCS(cS, "parentid", targetParentFolderID)
                                                        folder.save(cp)
                                                    End If
                                                    'Call main.CloseCS(cS)
                                                End If
                                                If Ptr = (targetFolderCnt - 1) Then
                                                    topFolderID = targetFolderId
                                                End If
                                            Next
                                        End If
                                    End If
                                End If
                            End If
                            Ptr += 1
                        Next
                        '
                    End If
                End If
                '
                LoadFolders_returnTopFolderId = topFolderID
                '
                result = LoadFolders_returnTopFolderId
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function

        'Private Function IsEmpty(folderCells As Object) As Boolean
        '    Throw New NotImplementedException()
        'End Function
        '
        '===============================================================================================
        '   returns the id of the cache folder that matches the target folder
        '       targetfolder = 'tier1\tier2\tier3'
        '       targetArray=['tier1','tier2','tier3'], targetArray(0)='tier1'
        '       targetArrayPtr is the index into the targetArray of the folder we are looking up
        '       returns the id of the folder 'tier3' that has a parent folder 'tier2', etc.
        '       if not folder exists, it returns 0
        '===============================================================================================
        '
        Private Function loadFolders_getFolderID(cp As CPBaseClass, targetArray() As String, targetArrayPtr As Integer) As Integer

            '
            Dim cachePtr As Integer
            Dim cacheFolderID As Integer
            Dim cacheParentFolderID As Integer
            Dim targetFolderName As String
            Dim targetFolderParentId As Integer
            '
            loadFolders_getFolderID = 0
            targetFolderName = targetArray(targetArrayPtr)
            cachePtr = FolderNameIndex.getPtr(targetFolderName)
            Do While cachePtr >= 0
                cacheFolderID = folders(cachePtr).FolderID
                If targetArrayPtr = 0 Then
                    '
                    ' this was the top-most folder, return the non-zero cache id
                    '
                    If folders(cachePtr).parentFolderID <> 0 Then
                        '
                        ' top of target path but record parent <> 0, try next record
                        '
                    Else
                        '
                        ' top of target path matches records (parentid=0)
                        '
                        loadFolders_getFolderID = cacheFolderID
                        Exit Do
                    End If
                Else
                    '
                    ' not top-most, since there could be multiple matches, test the parent
                    ' of this match, if it is ok, return with this id. If not, try the next
                    ' cache match for this folder name.
                    '
                    targetFolderParentId = loadFolders_getFolderID(cp, targetArray, targetArrayPtr - 1)
                    If targetFolderParentId <= 0 Then
                        '
                        ' parent folder not found, try the next matching folder name
                        '
                    Else
                        '
                        ' parent folder found, check that its target folder matches the parent id of this folder
                        '
                        If targetFolderParentId = folders(cachePtr).parentFolderID Then
                            '
                            ' this folder is correct, return with it's ID
                            '
                            loadFolders_getFolderID = cacheFolderID
                            Exit Do
                        Else
                            '
                            ' the cache folder hierarchy does not match the traget folder string, try next name match
                            '
                        End If
                    End If
                End If
                cachePtr = FolderNameIndex.GetNextPointerMatch(targetFolderName)
            Loop
            '
        End Function
        '
        '
        '
        Private Function LoadFolders_GetModifyAccess(cp As CPBaseClass, FolderID As Integer) As Boolean

            '
            Dim Ptr As Integer
            '
            Ptr = FolderIdIndex.getPtr(CStr(FolderID))
            If Ptr >= 0 Then
                If folders(Ptr).modifyAccessIsValid Then
                    '
                    '
                    '
                    LoadFolders_GetModifyAccess = folders(Ptr).hasModifyAccess
                ElseIf folders(Ptr).parentFolderID = 0 Then
                    '
                    ' Parent is root, this folder does not have access
                    '
                    LoadFolders_GetModifyAccess = False
                    folders(Ptr).hasModifyAccess = LoadFolders_GetModifyAccess
                    folders(Ptr).modifyAccessIsValid = True
                Else
                    '
                    ' Parent is not root
                    '
                    LoadFolders_GetModifyAccess = LoadFolders_GetModifyAccess(cp, folders(Ptr).parentFolderID)
                    folders(Ptr).hasModifyAccess = LoadFolders_GetModifyAccess
                    folders(Ptr).modifyAccessIsValid = True
                End If
            End If
            '
        End Function
        '
        '=================================================================================
        ' Create
        '=================================================================================
        '
        Private Sub Class_Initialize()
            iMinRows = 10
        End Sub
        '
        '=================================================================================
        ' Kill
        '=================================================================================
        '
        Private Sub Class_Terminate()
        End Sub
        '
        '
        '
        Private Function GetFileSize(cp As CPBaseClass, VirtualFilePathPage As String) As Integer
            Dim hint As String
            Dim result As Integer
            Try
                '
                hint = "1"
                '
                hint = "2"
                VirtualFilePathPage = Replace(VirtualFilePathPage, "/", "\")
                Dim SlashPosition As Integer = InStrRev(VirtualFilePathPage, "\")
                Dim Filename As String = ""
                Dim Pathname As String = ""
                If SlashPosition <> 0 Then
                    Filename = LCase(Mid(VirtualFilePathPage, SlashPosition + 1))
                    Pathname = Mid(VirtualFilePathPage, 1, SlashPosition - 1)
                End If
                '
                Dim FileDescriptor As String
                FileDescriptor = cp.File.fileList(Pathname)
                hint = "3"
                If FileDescriptor = "" Then
                    'Call AppendLogFile("GetFileSize, descriptor is blank")
                Else
                    hint = "4"
                    Dim FileSplit2() As String = Split(FileDescriptor, vbCrLf)
                    'Call AppendLogFile("GetFileSize, FileDescriptor lines=" & UBound(FileSplit2))
                    hint = "5"
                    Dim Ptr As Integer
                    For Ptr = 0 To UBound(FileSplit2)
                        Dim FileParts() As String = Split(FileSplit2(Ptr), vbTab)
                        If UBound(FileParts) <= 5 Then
                            'Call AppendLogFile("GetFileSize, FileDescriptor row [" & Ptr * "] has <6 parts, descrriptor=" & FileDescriptor)
                        Else
                            If LCase(FileParts(0)) = Filename Then
                                GetFileSize = cp.Utils.EncodeInteger(FileParts(5))
                                'Call AppendLogFile("GetFileSize, match on " & FileParts(0))
                                Exit For
                            End If
                        End If
                        result = GetFileSize
                    Next
                    hint = "6"
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return GetFileSize
        End Function
        '
        '
        '
        Private Function GetFileTypeID(cp As CPBaseClass, Filename As String) As Integer

            '
            Dim FileNameSplit() As String
            Dim FileExtension As String
            Dim CSType As Integer
            Dim DefaultFileTypeID As Integer
            Dim cnt As Integer
            Dim Ptr As Integer
            Dim hint As String
            '
            FileNameSplit = Split(Filename, ".")
            FileExtension = FileNameSplit(UBound(FileNameSplit))
            '
            ' try to read if from IconFiles
            '
            hint = "1"
            GetFileTypeID = DefaultFileTypeID
            For Each iconFile As FileType In iconFiles
                hint = "2"
                If InStr(1, "," & iconFile.ExtensionList & ",", "," & FileExtension & ",", vbTextCompare) <> 0 Then
                    hint = "3"
                    GetFileTypeID = iconFile.FileTypeID
                    Exit For
                    If LCase(iconFile.Name) = "default" Then
                        hint = "4"
                        DefaultFileTypeID = iconFile.FileTypeID
                    End If
                End If
            Next
            'cnt = IconFileCnt
            'If cnt > 0 Then
            '    For Ptr = 0 To cnt - 1
            '    Next
            'End If
            hint = "5"
            '    '
            '    ' try Db next
            '    '
            '    If GetFileTypeID = 0 Then
            'hint = "6"
            '        CSType = Main.OpenCSContent("Library File Types", "(extensionlist like '%," & FileExtension & ",%')or(extensionlist like '%,." & FileExtension & ",%')")
            '        If Main.IsCSOK(CSType) Then
            '            GetFileTypeID = Main.GetCSInteger(CSType, "ID")
            '        End If
            '        Call Main.closecs(CSType)
            '        If GetFileTypeID = 0 Then
            '            GetFileTypeID = Main.GetRecordID("Library File Types", "default")
            '        End If
            '    End If
            '
        End Function


    End Class
End Namespace
