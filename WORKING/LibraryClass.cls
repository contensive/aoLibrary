VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibraryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'==================================================================================================
'   Resource Library
'
'   Folder access
'       The resource library is configured to display files/folders from a provided topFolderPath folder down.
'
'       ViewAccess means you can view files and folders from the topFolderPath folder down
'           true for folders under the topFolderPath, otherwise false
'
'       You have ViewAccess to a folder if you are an administrator or if the folder is under the topFolderPath
'
'       If no topFolderPath is provided, the rootFolder is the topFolderPath
'
'       ModifyAccess means they can add folders and upload files from a folder down.
'
'       only administrators can have ModifyAccess on the root folder
'
'==================================================================================================
'
' ----- Global Objects
'
Private Main As Object
Private Csv As Object
'Private Main As ccWeb3.MainClass
Private Type FileType
    Name As String
    FileTypeID As Long
    ExtensionList As String
    IconFilename As String
    IsImage As Boolean
    IsFlash As Boolean
    IsVideo As Boolean
    MediaIconFilename As String
    IsDownload As Boolean
    DownloadIconFilename As String
End Type
Private IconFiles() As FileType
Private IconFileCnt As Long
'
'
'
Private Type FolderType
    FolderID As Long
    parentFolderID As Long
    Name As String
    FullPath As String
    '
    hasViewAccess As Boolean                    ' has permission to view this folder (below topFolderPath)
    viewAccessIsValid As Boolean                 ' true when hasViewAccess is correct
    '
    hasModifyAccess As Boolean                  ' has permission to modify files and folders in this folder
    modifyAccessIsValid As Boolean              ' true when hasModifyAccess is correct
End Type
Private folders() As FolderType
Private folderCnt As Long
Private FolderIdIndex As FastIndex5.FastIndex5Class
Private FolderNameIndex As FastIndex5.FastIndex5Class
Private FolderPathIndex As FastIndex5.FastIndex5Class
'
Private FolderSelect As String
'
' -----------------------------------------------------------------------------------
' ----- Publics
' -----------------------------------------------------------------------------------
' ----- not used
'
Private UserMemberID As Long
Private RequestStream As String
'
' ----- Icons used
'
Const IconFolderOpen = "<img src=""/ResourceLibrary/IconFolderOpen.gif"" border=""0"" width=""22"" height=""23"" ALT=""Close this folder"">"
Const IconFolderClosed = "<img src=""/ResourceLibrary/IconFolderClosed.gif"" border=""0"" width=""22"" height=""23"" alt=""Open this folder"">"
Const IconFolderAdd = "<img src=""/ResourceLibrary/IconFolderAdd2.gif"" border=""0"" width=""22"" height=""23"" alt=""Add a new folder"">"
Const IconFolderEdit = "<img src=""/ResourceLibrary/IconFolderEdit.gif"" border=""0"" width=""22"" height=""23"" alt=""Edit this folder"">"
Const IconFile = "<img src=""/ResourceLibrary/IconFile.gif"" border=""0"" width=""22"" height=""23"" alt=""file"">"
Const IconFileAdd = "<img src=""/ResourceLibrary/IconContentAdd.gif"" border=""0"" width=""18"" height=""22"" alt=""Add a new  file"">"
Const IconFileEdit = "<img src=""/ResourceLibrary/IconContentEdit.gif"" border=""0"" width=""18"" height=""22"" alt=""Edit this file"">"
Const IconPreview = "<img src=""/ResourceLibrary/IconPreview.gif"" border=""0"" width=""22"" height=""23"" alt=""Preview this image"">"
Const IconDownload = "<img src=""/ResourceLibrary/IconDownload3.gif"" border=""0"" width=""16"" height=""16"" alt=""Select this download"" valign=""absmiddle"">"
Const IconCreateImage = "<img src=""/ResourceLibrary/IconimagePlace.gif"" border=""0"" width=""18"" height=""22"" alt=""Select this image"">"
Const IconCreateDownload = "<img src=""/ResourceLibrary/IconDownload3.gif"" border=""0"" width=""16"" height=""16"" alt=""Select this download"" valign=""absmiddle"">"
Const IconSpacer = "<img src=""/ResourceLibrary/spacer.gif"" width=""22"" height=""23"">"
Const IconView = "<img src=""/ResourceLibrary/IconView.gif"" border=""0"" width=""22"" height=""23"" alt=""Preview this file"">"
Const IconImage = "<img src=""/ResourceLibrary/IconImage2.gif"" border=""0"" width=""22"" height=""23"" alt=""Image"">"
Const IconPDF = "<img src=""/ResourceLibrary/IconPDF.gif"" border=""0"" width=""16"" height=""16"" alt=""Adobe Pdf"">"
Const IconOther = "<img src=""/ResourceLibrary/IconFile.gif"" border=""0"" width=""22"" height=""23"" alt=""Unrecognized File Type"">"
Const IconNoFile = "<img src=/ResourceLibrary/BulletRound2.gif width=5 height=5>"
'
' ----- SelectResource Support
'       This means the resource library supports buttons that allow objects to be
'       placed on different page from the resource library, like an Editor
'
Private AllowPlace As Boolean
'
' ----- If an editor is used to call the resource library, the window.opener.insertresource()
'       call needs the object name of the editor so the contents can be copied to the invisible
'       form field after the changes (no onchange event available)
'
Private SelectResourceEditorObjectName As String
'
' ----- If AllowPlace is true and SelectLinkObjectName<>"", the RL is being used as a link selector
'       When the 'place' icon is clicked, the URL of the resource is copied to the window.opener.[selectlinkobjectname]
'
Private SelectLinkObjectName As String
'
' ----- Blocks the folder list in the left hand side
'
Private blockFolderNavigation As Boolean
'
' -----------------------------------------------------------------------------------
' ----- Privates
' -----------------------------------------------------------------------------------
'
Private iMinRows As Long
Private iFolderID  As Long                      ' Current Folder being Displayed, 0 for root
Private SourceMode As Long                      '
        '
'        ' SourceMode
'        '   3/6/2010 - moved codes up to capture the 0 case and set it to page
'        '   1 = From Editor Object or Link selector: allow image and download insert, provide close button
'        '   2 = From Editor Image Properties: allow image insert, provide close button
'        '   3 = From Admin site, no inserts, and provide cancel button
'Const SourceModeOnPage = 1
'Const SourceModeFromDownloadRequest = 2
'Const SourceModeFromLinkDialog = 3
        '   0 = From Editor Object selector: allow image and download insert, provide close button
        '   1 = From Editor Image Properties: allow image insert, provide close button
        '   2 = From Admin site, no inserts, and provide cancel button
Const SourceModeFromDownloadRequest = 0
Const SourceModeFromLinkDialog = 1
Const SourceModeOnPage = 2
        '
        '   0 caller is the editor directly, clicking on action icons calls InsertImaage, etc
        '   1 caller is the editor image page, clicking on action icons calls the image page methods
        '
Private HoldPosition As Long
'
'========================================================================
'   v3.3 Compatibility
'       To make an Add-on that works the same in v3.3 and v3.4, use this adapter instead of the execute above
'========================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Set Csv = CsvObject
    Call Init(MainObject)
    Execute = GetContent(OptionString)
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("Execute")
End Function
'
'=================================================================================
' Initilize this object
'   Must be called after creation, before any other calls
'=================================================================================
'
Public Sub Init(MainObject As Object)
    On Error GoTo ErrorTrap
    '
    Set Main = MainObject
    '
    Exit Sub
ErrorTrap:
    Call HandleRLTrapError("init")
End Sub
'
'=================================================================================
'   Aggregate Object Interface
'=================================================================================
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim topFolderPath As String
    Dim AllowGroupAdd As Boolean
    '
    topFolderPath = Main.getAddonOption("RootFolderName", OptionString)
    Call Main.testpoint("topFolderPath=[" & topFolderPath & "]")
    AllowGroupAdd = kmaEncodeBoolean(Main.getAddonOption("AllowGroupAdd", OptionString))
    AllowPlace = kmaEncodeBoolean(Main.getAddonOption("AllowSelectResource", OptionString))
    SelectResourceEditorObjectName = Main.getAddonOption("SelectResourceEditorObjectName", OptionString)
    SelectLinkObjectName = Main.getAddonOption("SelectLinkObjectName", OptionString)
    blockFolderNavigation = kmaEncodeBoolean(Main.getAddonOption("Block Folder Navigation", OptionString))
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
    GetContent = GetForm(topFolderPath, AllowGroupAdd)
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetForm")
End Function
'=================================================================================
' Returns the Resource Library HTML.
'   This HTML does not include the HTML, HEAD or BODY tags.
'=================================================================================
'
Private Function GetResourceLibrary() As String
    GetResourceLibrary = GetForm("", False)
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetResourceLibrary")
End Function
'
'=================================================================================
' Returns the Resource Library HTML.
'   This HTML does not include the HTML, HEAD or BODY tags.
'=================================================================================
'
Private Function GetForm(topFolderPath As String, AllowGroupAdd As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Const LibraryFileTypespathFilename = "ResourceLibrary\LibraryConfig.xml"
    '
    Dim currentFolderPtr As Long
    Dim currentFolderID As Long
    Dim currentFolderHasModifyAccess As Boolean
    '
    Dim reloadFolderCache As Boolean
    Dim IsContentManagerFiles As Boolean
    Dim IsContentManagerFolders As Boolean
    Dim AltSizes() As String
    Dim AltSize As String
    Dim AltSizeHeight As Long
    Dim BestFitHeight As Long
    Dim BestFitWidth As Long
    Dim BestFitAltSize As String
    Dim ThumbNailSrc As String
    Dim FieldList As String
    Dim JumpSelect As String
    Dim VirtualFilePathPage As String
    Dim Pos As Long
    Dim ImageFilename As String
    Dim doc As DOMDocument60
    Dim node As IXMLDOMElement
    Dim baseNode As IXMLDOMElement
    Dim typeNode As IXMLDOMElement
    Dim FilePath As String
    Dim IconCnt As Long
    Dim Filename As String
    Dim ColumnCnt As Long
    Dim AllowPlaceColumn As Boolean
    Dim AllowEditColumn As Boolean
    Dim AllowSelectColumn As Boolean
    Dim BuildVersion As String
    Dim AltSizeList As String
    Dim FilenameNoExtension As String
    Dim sf As SfImageResize.ImageResize
    Dim SQL As String
    Dim UpdateRecord As Boolean
    Dim FormFolders As String
    Dim FormDetails As String
    Dim FileExtension As String
    Dim FileNameSplit() As String
    Dim CSType As Long
    Dim FileTypeID As Long
    Dim FileTypeFilter As String
    Dim RowValues As Variant
    Dim RowPtr As Long
    Dim RowCnt As Long
    Dim MoveFolderID As Long
    Dim MoveFileID As Long
    Dim targetFolderId As Long
    Dim IconName As String
    Dim DownloadName As String
    Dim MediaName As String
    Dim DefaultIcon As String
    Dim DefaultMedia As String
    Dim DefaultDownload As String
    Dim UseDefaults As Boolean
    Dim cS As Long
    Dim RowCount As Long
    Dim SortField As String
    Dim SortDirection As Long
    Dim Criteria As String
    Dim ChildFolderID As Long
    Dim ChildName As String
    Dim ChildFolderName As String
    'Dim AllowUpFolder As Boolean
    Dim ImageSrc As String
    Dim IconLink As String
    Dim IconOnClick  As String
    Dim EditLink As String
    Dim ModifiedDate As Date
    Dim Description As String
    Dim RecordName As String
    Dim ImageAlt As String
    Dim ImageWidth As Long
    Dim ImageHeight As Long
    Dim ImageWidthText As String
    Dim ImageHeightText As String
    Dim ResourceRecordID As Long
    Dim ResourceHref As String
    Dim DotPosition As Long
    Dim AddFolderEditLink As String
    Dim AllowFolderAuthoring As Boolean
    Dim AllowFileAuthoring As Boolean
    Dim FolderCID As Long
    Dim FileCID As Long
    Dim ParentFolderName As String
    Dim parentFolderID As Long
    Dim RowName As String
    Dim RowFeatures As String
    Dim RowDescription As String
    Dim VirtualFilePath As String
    Dim ConfigFilename As String
    Dim topFolderID As Long
    'Dim FolderGroupName As String
    Dim FolderParentID As Long
    'Dim FolderGroupID As Long
    Dim AllowLocalFileAdd As Boolean
    Dim ButtonBar As String
    Dim Button As String
    Dim UploadCount As Long
    Dim UploadPointer As Long
    Dim Copy As String
    Dim ButtonBarStyle As String
    Dim OptionPanelStyle As String
    Dim AllowThumbnails As Boolean
    Dim FolderIDString As String
    Dim Link As String
    Dim DeleteFolderID As Long
    Dim DeleteFileID As Long
    Dim Ptr As Long
    Dim FileID As Long
    Dim folderName As String
    Dim fileSize As Long
    Dim Pathname As String
    Dim SlashPosition As Long
    Dim FileDescriptor As String
    Dim FileSplit As String
    Dim FileSplit2() As String
    Dim FileParts() As String
    Dim FileCount As Long
    Dim ButtonExit
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
    ButtonBarStyle = "" _
        & " color: black;" _
        & " font-weight: bold;" _
        & " padding: 5px;" _
        & " background-color: #a0a0a0;" _
        & " border-bottom: 1px solid #e0e0e0;" _
        & " border-right: 1px solid #e0e0e0;" _
        & " border-top: 1px solid #808080;" _
        & " border-left: 1px solid #808080;"
    '
    OptionPanelStyle = "" _
        & " color: black;" _
        & " font-weight: bold;" _
        & " padding: 5px;" _
        & " background-color: #d0d0d0;" _
        & " border-bottom: 1px solid #e0e0e0;" _
        & " border-right: 1px solid #e0e0e0;" _
        & " border-top: 1px solid #a0a0a0;" _
        & " border-left: 1px solid #a0a0a0;"
    '
    If Not (Main Is Nothing) Then
        '
        ' Determine Current Folder
        '
        hint = "001"
        BuildVersion = Main.SiteProperty_BuildVersion
        IsContentManagerFiles = Main.IsContentManager("Library Files")
        IsContentManagerFolders = Main.IsContentManager("Library Folders")
        Button = Main.GetStreamText("Button")
        FileTypeFilter = LCase(Main.GetStreamText("ffilter"))
        Call Main.AddRefreshQueryString("ffilter", FileTypeFilter)
        AllowThumbnails = Main.GetMemberProperty("LibraryAllowthumbnails", False)
        FolderIDString = Main.GetStreamText("folderid")
        currentFolderID = Main.EncodeInteger(FolderIDString)
        If FolderIDString <> "" Then
            Call Main.SetMemberProperty("LibraryFolderID", currentFolderID)
        Else
            currentFolderID = Main.GetMemberProperty("Libraryfolderid", 0)
        End If
        '
        ' Load Folder cache
        '
        hint = "010, topFolderPath=" & topFolderPath
        topFolderID = LoadFolders_returnTopFolderId(topFolderPath)
        reloadFolderCache = False
        '
        ' verify that current folder has viewAccess (if not jumpt to root)
        '
        If currentFolderID <> 0 Then
            currentFolderPtr = FolderIdIndex.GetPointer(CStr(currentFolderID))
            If (currentFolderPtr > UBound(folders)) Or (currentFolderPtr < 0) Then
                currentFolderPtr = 0
            End If
            If currentFolderID < 0 Then
                currentFolderID = 0
                Call Main.SetMemberProperty("LibraryFolderID", currentFolderID)
            ElseIf Not folders(currentFolderPtr).hasViewAccess Then
                currentFolderID = 0
                Call Main.SetMemberProperty("LibraryFolderID", currentFolderID)
            End If
        End If
        '
        ' determine if current folder has modify access
        '
        hint = "020"
        currentFolderHasModifyAccess = False
        If (Main.IsAdmin Or IsContentManagerFiles Or IsContentManagerFolders) Then
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
            currentFolderPtr = FolderIdIndex.GetPointer(CStr(currentFolderID))
            If currentFolderPtr >= 0 Then
                currentFolderHasModifyAccess = folders(currentFolderPtr).hasModifyAccess
            End If
        End If
        'topFolderID = GetFolderID(topFolderPath)
        '
        ' Load IconFiles
        '
        hint = "030"
        Set doc = New DOMDocument60
        FilePath = Main.serverprotocol & Main.ServerHost & Main.serverFilePath
        ConfigFilename = Main.physicalFilePath & LibraryFileTypespathFilename
        Call doc.Load(ConfigFilename)
        If doc.parseError.errorCode <> 0 Then
            '
            ' Error
            '
            'Call AppendLogFile2( "Server", "AddonInstallClass", "DownloadCollectionFiles, The GetCollection request for GUID [" & CollectionGuid & "] failed. The error was [" & doc.parseError.reason & "]")
        Else
            hint = "040"
            If (LCase(doc.documentElement.baseName) <> LCase("libraryconfig")) Then
                'Return_ErrorMessage = "The collection file from the server was not valid for collection [" & CollectionGuid & "]"
                'DownloadCollectionFiles = False
                'Call AppendClassLogFile("Server", "AddonInstallClass", "DownloadCollectionFiles, The GetCollection request for GUID [" & CollectionGuid & "] named [" & Collectionname & "] returned a file with a bad format. The root node was [" & doc.documentElement.baseName & "] but [" & DownloadFileRootNode & "] was expected.")
            Else
                If doc.documentElement.childNodes.Length = 0 Then
                    'Return_ErrorMessage = "The collection file from the server was empty for collection [" & CollectionGuid & "]"
                    'Call AppendClassLogFile("Server", "AddonInstallClass", "DownloadCollectionFiles, The GetCollection request for GUID [" & CollectionGuid & "] named [" & Collectionname & "] returned a file with no nodes. The collection was probably not found")
                    'DownloadCollectionFiles = False
                Else
                    With doc.documentElement
                        Ptr = 0
                        hint = "050"
                        For Each baseNode In .childNodes
                            hint = "060"
                            Select Case LCase(baseNode.baseName)
                                Case "filetype"
                                    hint = "070"
                                    Ptr = Ptr + 1
                                    If Ptr >= IconCnt Then
                                        IconCnt = IconCnt + 10
                                        ReDim Preserve IconFiles(IconCnt)
                                    End If
                                    With IconFiles(Ptr)
                                    For Each typeNode In baseNode.childNodes
                                        Select Case LCase(typeNode.baseName)
                                            Case "name"
                                                .Name = typeNode.Text
                                            Case "filetypeid"
                                                .FileTypeID = kmaEncodeInteger(typeNode.Text)
                                            Case "extensionlist"
                                                .ExtensionList = typeNode.Text
                                            Case "isdownload"
                                                .IsDownload = kmaEncodeBoolean(typeNode.Text)
                                            Case "isimage"
                                                .IsImage = kmaEncodeBoolean(typeNode.Text)
                                            Case "isvideo"
                                                .IsVideo = kmaEncodeBoolean(typeNode.Text)
                                            Case "isflash"
                                                .IsFlash = kmaEncodeBoolean(typeNode.Text)
                                            Case "iconlink"
                                                .IconFilename = typeNode.Text
                                            Case "mediaiconlink"
                                                .MediaIconFilename = typeNode.Text
                                            Case "downloadiconlink"
                                                .DownloadIconFilename = typeNode.Text
                                        End Select
                                    Next
                                End With
                            End Select
                        Next
                    IconFileCnt = Ptr
                    End With
                End If
            End If
        End If
        '
        ' Verify default icons
        '
        hint = "100"
        DefaultIcon = Main.GetDefaultValue("ResourceLibraryIconDefault")
        If DefaultIcon = "" Then
            DefaultIcon = "\cclib\images\IconImage2.gif"
        End If
        DefaultMedia = Main.GetDefaultValue("ResourceLibraryIconDefaultMedia")
        If DefaultMedia = "" Then
            DefaultMedia = "\cclib\images\Iconimage2Media.gif"
        End If
        DefaultDownload = Main.GetDefaultValue("ResourceLibraryIconDefaultDownload")
        If DefaultDownload = "" Then
            DefaultDownload = "\cclib\images\Iconimage2Download.gif"
        End If
        '
        If Main.GetStreamText("SourceMode") = "" Then
            SourceMode = SourceModeOnPage
        Else
            SourceMode = Main.GetStreamInteger("SourceMode")
        End If
        ' 3/6/2010 - changed SourceMode constants so 0 is invalid (default), set to page
        'If SourceMode = 0 Then
        '    SourceMode = SourceModeOnPage
        'End If
        Call Main.AddRefreshQueryString("SourceMode", SourceMode)
        '
        ' ----- verify currentFolderID
        '
        If currentFolderID = 0 Then
            '
            ' No folder give, use root folder, no owner
            '
            currentFolderID = topFolderID
            Call Main.SetMemberProperty("LibraryFolderID", currentFolderID)
            'FolderGroupName = ""
        End If
        If currentFolderID <> 0 Then
            cS = Main.OpenCSContentRecord("Library Folders", currentFolderID, , , "ParentID")
            'CS = Main.OpenCSContentRecord("Library Folders", FolderID, , , "ParentID,GroupID")
            If Main.IsCSOK(cS) Then
                FolderParentID = Main.GetCSInteger(cS, "ParentID")
            End If
            Call Main.closecs(cS)
        End If
        If (topFolderID <> currentFolderID) And (topFolderID <> FolderParentID) Then
            '
            ' Check if Folder is under the given root folder
            '
            If Not IsInFolder(topFolderID, currentFolderID) Then
                '
                ' Current folder is not in Root Folder, Use Root Folder
                '
                currentFolderID = topFolderID
                Call Main.SetMemberProperty("LibraryFolderID", currentFolderID)
            End If
        End If
        '
        ' ----- Determine if select, place or edit icons are allowed
        '
        ColumnCnt = 5
        AllowPlaceColumn = AllowPlace And ((SourceMode = SourceModeFromLinkDialog) Or (SourceMode = SourceModeFromDownloadRequest))
        'AllowPlaceColumn = AllowPlace And (IsContentManagerFiles Or IsContentManagerFolders) And ((SourceMode = SourceModeFromLinkDialog) Or (SourceMode = SourceModeFromDownloadRequest))
        If AllowPlaceColumn Then
            ColumnCnt = ColumnCnt + 1
        'Else
        '    ColumnCnt = ColumnCnt + 1
        End If
        AllowEditColumn = (IsContentManagerFiles Or IsContentManagerFolders)
        If AllowEditColumn Then
            ColumnCnt = ColumnCnt + 1
        End If
        AllowSelectColumn = currentFolderHasModifyAccess
        If AllowSelectColumn Then
            ColumnCnt = ColumnCnt + 1
        End If
        '
        ' ----- Setup folder editing
        '
        AllowFolderAuthoring = IsContentManagerFolders
        If AllowFolderAuthoring Then
            FolderCID = Main.GetContentID("Library Folders")
        End If
        '
        ' ----- Setup file editing
        '
        AllowFileAuthoring = IsContentManagerFiles
        If AllowFileAuthoring Then
            FileCID = Main.GetContentID("Library Files")
        End If
        '
        ' ----- Setup Local File Management
        '       Allow if Content Manager or user has group membership
        '       Always allow, everyone has access to the root folder, then if you can get to the folder, let em upload
        '
        AllowLocalFileAdd = True
        'AllowLocalFileAdd = AllowFileAuthoring Or (AllowGroupAdd And Main.IsAdmin)
        'AllowLocalFileAdd = AllowFileAuthoring Or (AllowGroupAdd And (Main.IsAdmin Or Main.IsGroupMember(FolderGroupName)))
        '
        ' ----- Process input
        '
        hint = "200"
        If Button <> "" Then
            AllowThumbnails = Main.GetStreamBoolean("AllowThumbnails")
            Call Main.SetMemberProperty("LibraryAllowthumbnails", AllowThumbnails)
        Select Case Button
            Case ButtonCancel
                '
                ' CAncel button, just redirect back to the current page
                '
                Link = Main.ServerPage
                'If Main.RefreshQueryString <> "" Then
                '    Link = Link & "?" & Main.RefreshQueryString
                'End If
                Call Main.Redirect(Link)
            Case ButtonDelete
                '
                '
                '
                RowCount = Main.GetStreamInteger("RowCount")
                If RowCount > 0 Then
                    For Ptr = 0 To RowCount - 1
                        If Main.GetStreamBoolean("Row" & Ptr) Then
                            DeleteFolderID = Main.GetStreamInteger("Row" & Ptr & "FolderID")
                            If DeleteFolderID <> 0 Then
                                'Call Main.WriteStream("Deleting Folder " & FolderID)
                                Call Main.DeleteContentRecord("Library Folders", DeleteFolderID)
                                reloadFolderCache = True
                            End If
                            DeleteFileID = Main.GetStreamInteger("Row" & Ptr & "FileID")
                            If DeleteFileID <> 0 Then
                                'Call Main.WriteStream("Deleting File " & FileID)
                                Call Main.DeleteContentRecord("Library Files", DeleteFileID)
                                reloadFolderCache = True
                            End If
                        End If
                    Next
                End If
            Case ButtonApply
                '
                ' Move Files
                '
                If Main.GetStreamBoolean("Move") Then
                    targetFolderId = Main.GetStreamInteger("MoveFolderID")
                    RowCount = Main.GetStreamInteger("RowCount")
                    If RowCount > 0 Then
                        For Ptr = 0 To RowCount - 1
                            If Main.GetStreamBoolean("Row" & Ptr) Then
                                MoveFolderID = Main.GetStreamInteger("Row" & Ptr & "FolderID")
                                MoveFileID = Main.GetStreamInteger("Row" & Ptr & "FileID")
                                If MoveFolderID <> 0 Then
                                    Call Main.executeSQL("default", "update ccLibraryFolders set ParentID=" & targetFolderId & " where ID=" & MoveFolderID)
                                    reloadFolderCache = True
                                ElseIf MoveFileID <> 0 Then
                                    Call Main.executeSQL("default", "update ccLibraryFiles set FolderID=" & targetFolderId & " where ID=" & MoveFileID)
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
                    UploadCount = Main.GetStreamInteger("AddFolderCount")
                    For UploadPointer = 1 To UploadCount
                        folderName = Main.GetStreamText("FolderName." & UploadPointer)
                        If folderName <> "" Then
                            If IsContentManagerFolders And (Not Main.IsAdmin) And (currentFolderID = 0) Then
                                '
                                ' Content Managers can not add folders to the root folder
                                '
                                Call Main.AddUserError("Your account does not have access to add new folders to the root folder.")
                                Exit For
                            Else
                                cS = Main.InsertCSRecord("Library Folders")
                                If Main.IsCSOK(cS) Then
                                    Copy = Main.GetStreamText("FolderDescription." & UploadPointer)
                                    Call Main.SetCS(cS, "Name", folderName)
                                    Call Main.SetCS(cS, "Description", Copy)
                                    If currentFolderID <> 0 Then
                                        Call Main.SetCS(cS, "ParentID", currentFolderID)
                                    End If
                                End If
                                Call Main.closecs(cS)
                                reloadFolderCache = True
                            End If
                        End If
                    Next
                    '
                    ' Upload files
                    '
                    hint = "400"
                    UploadCount = Main.GetStreamInteger("LibraryUploadCount")
                    For UploadPointer = 1 To UploadCount
                        ImageFilename = Main.GetStreamText(RequestNameLibraryUpload & "." & UploadPointer)
                        If ImageFilename <> "" Then
                            hint = "410"
                            cS = Main.InsertCSRecord("Library Files")
                            If Main.IsCSOK(cS) Then
                                Copy = Main.GetStreamText(RequestNameLibraryName & "." & UploadPointer)
                                If Copy = "" Then
                                    Copy = ImageFilename
                                End If
                                Call Main.SetCS(cS, "Name", Copy)
                                Copy = Main.GetStreamText(RequestNameLibraryDescription & "." & UploadPointer)
                                If Copy = "" Then
                                    Copy = ImageFilename
                                End If
                                FileExtension = ""
                                FilenameNoExtension = ""
                                AltSizeList = ""
                                Pos = InStrRev(ImageFilename, ".")
                                If Pos > 0 Then
                                    FileExtension = Mid(ImageFilename, Pos + 1)
                                    FilenameNoExtension = Left(ImageFilename, Pos - 1)
                                End If
                                hint = "420"
                                VirtualFilePathPage = Main.GetCSFilename(cS, "Filename", ImageFilename, "Library Files")
                                VirtualFilePath = Replace(VirtualFilePathPage, ImageFilename, "")
                                Call Main.SetCS(cS, "Description", Copy)
                                Call Main.SetCS(cS, "FolderID", currentFolderID)
                                Call Main.ProcessFormInputFile(RequestNameLibraryUpload & "." & UploadPointer, VirtualFilePath)
                                Call Main.SetCS(cS, "Filesize", GetFileSize(Main.physicalFilePath & VirtualFilePathPage))
Dim FileType As String
                                hint = "425"
                                FileTypeID = GetFileTypeID(ImageFilename)
                                Call Main.SetCS(cS, "FileTypeID", FileTypeID)
                                hint = "430"
                                If IconFiles(FileTypeID).IsImage And (BuildVersion > "3.4.190") Then
                                    hint = "440"
                                    '
                                    ' add image resize values
                                    '
                                    Set sf = CreateObject("sfimageresize.imageresize")
                                    sf.Algorithm = 5
                                    On Error Resume Next
                                    sf.LoadFromFile (Main.physicalFilePath & VirtualFilePathPage)
                                    If Err.Number = 0 Then
                                        ImageWidth = sf.Width
                                        ImageHeight = sf.Height
                                        Call Main.SetCS(cS, "height", ImageHeight)
                                        Call Main.SetCS(cS, "width", ImageWidth)
                                    Else
                                        Err.Clear
                                    End If
                                    '
                                    ' Attempt to make 640x
                                    '
                                    If sf.Width >= 640 Then
                                        sf.Width = 640
                                        Call sf.DoResize
                                        Call sf.SaveToFile(Main.physicalFilePath & VirtualFilePath & FilenameNoExtension & "-640x" & sf.Height & "." & FileExtension)
                                        AltSizeList = AltSizeList & vbCrLf & "640x" & sf.Height
                                    End If
                                    '
                                    ' Attempt to make 320x
                                    '
                                    If sf.Width >= 320 Then
                                        sf.Width = 320
                                        Call sf.DoResize
                                        Call sf.SaveToFile(Main.physicalFilePath & VirtualFilePath & FilenameNoExtension & "-320x" & sf.Height & "." & FileExtension)
                                        AltSizeList = AltSizeList & vbCrLf & "320x" & sf.Height
                                    End If
                                    '
                                    ' Attempt to make 160x
                                    '
                                    If sf.Width >= 160 Then
                                        sf.Width = 160
                                        Call sf.DoResize
                                        Call sf.SaveToFile(Main.physicalFilePath & VirtualFilePath & FilenameNoExtension & "-160x" & sf.Height & "." & FileExtension)
                                        AltSizeList = AltSizeList & vbCrLf & "160x" & sf.Height
                                    End If
                                    Set sf = Nothing
                                    Call Main.SetCS(cS, "AltSizeList", AltSizeList)
                                End If
                            End If
                            Call Main.closecs(cS)
                            reloadFolderCache = True
                        End If
                    Next
                End If
        
        End Select
        End If
        hint = "500"
        If reloadFolderCache Then
            folderCnt = 0
            topFolderID = LoadFolders_returnTopFolderId(topFolderPath)
            reloadFolderCache = False
        End If
        '
        ' Housekeep potential issue where a parent is deleted and child does not show
        '
        SQL = "update cclibraryfolders" _
            & " Set parentid=null" _
            & " where id in" _
            & " (" _
            & " select c.id from (cclibraryfolders c left join cclibraryfolders p on p.id=c.parentid)" _
            & " where p.ID Is Null" _
            & " and c.parentid is not null" _
            & " and c.parentid<>0" _
            & " )"
        Call Main.executeSQL("default", SQL)
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
        Call Main.executeSQL("default", SQL)
        '
        ' ----- Begin output
        '
        If (SourceMode = SourceModeFromDownloadRequest) Or (SourceMode = SourceModeFromLinkDialog) Then
            ButtonExit = Main.GetFormButton(ButtonClose, , , "window.close();")
        Else
            ButtonExit = Main.GetFormButton(ButtonCancel)
        End If
        If AllowLocalFileAdd Then
            If currentFolderHasModifyAccess Then
            'If IsContentManagerFiles Or IsContentManagerFolders Then
                ' <div style=""" & ButtonBarStyle & """>
                ButtonBar = "<div class=ccAdminButtonBar>" _
                    & ButtonExit _
                    & Main.GetFormButton(ButtonApply) _
                    & Main.GetFormButton(ButtonDelete, RequestNameButton, , "DeleteCheck();") _
                    & "</div>"
            Else
                ButtonBar = "<div class=ccAdminButtonBar>" _
                    & Main.GetFormButton(ButtonApply) _
                    & "</div>"
            End If
        End If
        GetForm = GetForm & Main.GetUploadFormStart() _
            & Main.GetFormInputHidden("FolderID", currentFolderID) _
            & Main.GetFormInputHidden(RequestNameRefreshBlock, Main.GetFormSN) _
            & ButtonBar _
            & ""
        'If Main.IsAdmin Then
        hint = "600"
        JumpSelect = GetJumpFolderPathSelect(currentFolderID, topFolderPath)
        GetForm = GetForm & "<div style=""padding:10px;"">" & GetParentFoldersLink(topFolderPath, topFolderID, currentFolderID, currentFolderID, Main.RefreshQueryString, "") & "</div>"
        If JumpSelect <> "" Then
            GetForm = GetForm & "<div style=""padding:10px;padding-top:0px"">" & "Jump to&nbsp;" & JumpSelect & "</div>"
        End If
        '
        ' From here down the form divides into FormFolder and FormDetails
        '
        FormDetails = FormDetails _
            & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" _
            & "<tr class=""headRow"">"
        If AllowSelectColumn Then
            FormDetails = FormDetails _
                & GetForm_HeaderCell("center", "10", "Select<BR>" & Image10)
        End If
        If AllowEditColumn Then
            FormDetails = FormDetails _
                & GetForm_HeaderCell("center", "15", "Edit<br>" & Image15)
        End If
        If AllowPlaceColumn Then
            FormDetails = FormDetails _
                & GetForm_HeaderCell("center", "15", "Place<br>" & Image15)
        'Else
        '    FormDetails = FormDetails _
        '        & GetForm_HeaderCell("center", "15", "&nbsp;<br>" & Image15)
        End If
        FormDetails = FormDetails _
            & GetForm_HeaderCell("left", "20", "&nbsp;<BR>" & Image20) _
            & GetForm_HeaderCell("left", "20%", "Name<br>" & Image20) _
            & GetForm_HeaderCell("left", "50%", "Description<br>" & Image15) _
            & GetForm_HeaderCell("center", "50", "Size<br>" & Image50) _
            & GetForm_HeaderCell("center", "50", "Modified&nbsp;&nbsp;<br>" & Image50) _
            & "</tr>"
        '
        ' ----- Select the Folder Rows
        '
        Criteria = "((ParentID is null)or(ParentID=0))"
        '
        If currentFolderID <> 0 Then
            Call Main.AddRefreshQueryString("FolderID", currentFolderID)
        End If
        '
        SortField = Main.GetStreamText("sortfield")
        If SortField = "" Then
            SortField = "Name"
        End If
        Call Main.AddRefreshQueryString("SortField", SortField)
        '
        SortDirection = Main.GetStreamInteger("sortdirection")
        If SortDirection <> 0 Then
            Call Main.AddRefreshQueryString("SortDirection", SortDirection)
        End If
        '
        If SortDirection <> 0 And SortField <> "" Then
            SortField = SortField & " DESC"
        End If
        '
        If currentFolderID <> 0 Then
            '
            ' ----- FolderID given, lookup record and get ParentID
            '       Note that allowupfolder allows users to "up" past top if they set it manually
            '       Fix this when security is added
            '
            Criteria = "(ID=" & KmaEncodeSQLNumber(currentFolderID) & ")"
            cS = Main.OpenCSContent_Internal("Library Folders", Criteria, , , , , "ParentID")
            If Main.IsCSOK(cS) Then
                parentFolderID = kmaEncodeInteger(Main.GetCSField(cS, "ParentID"))
                ParentFolderName = Main.GetCSEncodedField(cS, "Name")
            End If
            Call Main.closecs(cS)
            Criteria = "(ParentID=" & KmaEncodeSQLNumber(currentFolderID) & ")"
        ElseIf topFolderPath <> "" Then
            '
            ' ----- Rootfolder given, lookup record and get ParentID
            '
            ParentFolderName = topFolderPath
            Criteria = "(name=" & KmaEncodeSQLText(topFolderPath) & ")"
            cS = Main.OpenCSContent_Internal("Library Folders", Criteria, , , , , "ID")
            If Main.IsCSOK(cS) Then
                parentFolderID = 0
                currentFolderID = Main.GetCSInteger(cS, "ID")
                Call Main.SetMemberProperty("LibraryFolderID", currentFolderID)
            End If
            Call Main.closecs(cS)
            Criteria = "(ParentID=" & KmaEncodeSQLNumber(currentFolderID) & ")"
        Else
            '
            ' ----- Use Root as top (no record)
            '
        End If
        '
        ' ----- Output the page
        '
        RowCount = 0
        hint = "700"
        If True Then
            '
            ' check for viewAccess done at top.
            '
'        If Not hasModifyAccessByFolder(FolderID, topFolderPath) Then
'            '
'            ' ----- No folder access
'            '
'            FormDetails = FormDetails & "<tr class=""listRow""><td class=""left"">" & IconSpacer & "</td><td class=""left"" COLSPAN=" & ColumnCnt - 1 & ">You do not have access to this folder</td></tr>"
'            RowCount = RowCount + 1
'        Else
            '
            ' ----- List out the folders
            '
            cS = Main.OpenCSContent_Internal("Library Folders", Criteria, SortField, False, , , "ID,Name,ModifiedDate,DateAdded,Description,ParentID")
            If Main.IsCSOK(cS) Then
                Do While Main.IsCSOK(cS)
                    ChildFolderID = kmaEncodeInteger(Main.GetCSField(cS, "ID"))
                    '
                    ' currentFolder checked for viewAccess, so child folders automatically have viewAccess
                    '
                    FolderAccess = True
                    'If currentFolderID = 0 Then
                    '    '
                    '    ' Root folder, only show the folders which this person has hass
                    '    '
                    '    FolderAccess = hasModifyAccessByFolder(ChildFolderID, topFolderPath)
                    'Else
                    '    '
                    '    ' other folders - if you have access to current folder, you have access to all the rest
                    '    '
                    '    FolderAccess = True
                    'End If
                    If FolderAccess Then
                        ChildFolderName = kmaEncodeText(Main.GetCSField(cS, "Name"))
                        If ChildFolderName = "" Then
                            ChildFolderName = "[no name]"
                        End If
                        If AllowFolderAuthoring Then
                            EditLink = Main.SiteProperty_AdminURL & "?cid=" & FolderCID & "&id=" & ChildFolderID & "&af=4" & "&aa=2&depth=1"
                        Else
                            EditLink = ""
                        End If
                        IconLink = ModifyQueryString(Main.RefreshQueryString, "folderid", CStr(ChildFolderID))
                        ModifiedDate = KmaEncodeDate(Main.GetCSField(cS, "ModifiedDate"))
                        If ModifiedDate = 0 Then
                            ModifiedDate = KmaEncodeDate(Main.GetCSField(cS, "DateAdded"))
                        End If
                        Description = kmaEncodeText(Main.GetCSField(cS, "Description"))
                        FormDetails = FormDetails & GetFormRow_ChildFolders(IconFolderClosed, IconLink, "", ChildFolderName, "", ModifiedDate, RowCount, EditLink, Description, "CHILD", "", "", "", "", "", 0, ChildFolderID, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn)
                        RowCount = RowCount + 1
                    End If
                    Main.NextCSRecord (cS)
                    Loop
                End If
            Call Main.closecs(cS)
            '
            ' Lookup the files in the folder
            '
            hint = "720"
            If currentFolderID = 0 Then
                Criteria = "((FolderID is null)or(FolderID=0))"
            Else
                Criteria = "(FolderID=" & KmaEncodeSQLNumber(currentFolderID) & ")"
            End If
            If BuildVersion < "4.0.000" Then
                FieldList = "ID,Name,ModifiedDate,Filename,Width,Height,DateAdded,Description,AltText,FileTypeID,FileSize,'' as AltSizeList"
            Else
                FieldList = "ID,Name,ModifiedDate,Filename,Width,Height,DateAdded,Description,AltText,FileTypeID,FileSize,AltSizeList"
            End If
            cS = Main.OpenCSContent_Internal("Library Files", Criteria, SortField, False, , , FieldList)
            '
            ' List out the FILES
            '
            If Main.IsCSOK(cS) Then
                RowValues = Main.GetCSRows(cS)
                RowCnt = UBound(RowValues, 2) + 1
                For RowPtr = 0 To RowCnt - 1
                    UpdateRecord = False
                    ResourceRecordID = kmaEncodeInteger(RowValues(0, RowPtr))
                    RecordName = kmaEncodeText(RowValues(1, RowPtr))
                    ModifiedDate = KmaEncodeDate(RowValues(2, RowPtr))
                    Filename = kmaEncodeText(RowValues(3, RowPtr))
                    ImageWidthText = kmaEncodeText(RowValues(4, RowPtr))
                    ImageHeightText = kmaEncodeText(RowValues(5, RowPtr))
                    If ModifiedDate = CDate(0) Then
                        ModifiedDate = KmaEncodeDate(RowValues(6, RowPtr))
                    End If
                    Description = kmaEncodeText(RowValues(7, RowPtr))
                    ImageAlt = kmaEncodeText(RowValues(8, RowPtr))
                    FileTypeID = kmaEncodeInteger(RowValues(9, RowPtr))
                    fileSize = kmaEncodeInteger(RowValues(10, RowPtr))
                    AltSizeList = kmaEncodeText(RowValues(11, RowPtr))
                    '
                    ImageSrc = Main.serverFilePath & Replace(Filename, "\", "/")
                    '
                    DotPosition = InStrRev(ImageSrc, ".")
                    If DotPosition = 0 Then
                        FileExtension = ""
                        FilenameNoExtension = ""
                    Else
                        FileExtension = UCase(Mid(ImageSrc, DotPosition + 1))
                        FilenameNoExtension = Mid(ImageSrc, 1, DotPosition - 1)
                    End If
                    '
                    If FileTypeID = 0 Then
                        FileTypeID = GetFileTypeID(ImageSrc)
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
                        Call Main.SetCS(cS, "name", RecordName)
                    End If
                    '
                    ResourceHref = ""
                    IconLink = ""
                    If AllowFileAuthoring Then
                        EditLink = Main.SiteProperty_AdminURL & "?cid=" & FileCID & "&id=" & ResourceRecordID & "&af=4" & "&aa=2&depth=1"
                    Else
                        EditLink = ""
                    End If
                    '
                    ' create thumbnail
                    '
                    If AllowThumbnails Then
                        ThumbNailSrc = ImageSrc
                        If (FilenameNoExtension <> "") And (AltSizeList <> "") Then
                            AltSizes = Split(AltSizeList, vbCrLf)
                            BestFitHeight = 9999999
                            BestFitAltSize = ""
                            For Ptr = 0 To UBound(AltSizes)
                                '
                                ' Find the smallest image larger then height 50
                                '
                                AltSize = Trim(AltSizes(Ptr))
                                If AltSize <> "" Then
                                Pos = InStr(AltSize, "x")
                                If Pos > 0 Then
                                    AltSizeHeight = kmaEncodeInteger(Mid(AltSize, Pos + 1))
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
                        Pathname = Main.physicalFilePath & Replace(Filename, "/", "\")
                        fileSize = GetFileSize(Pathname)
                        If fileSize <> 0 Then
                            UpdateRecord = True
                        End If
                    End If
                    '
                    '
                    '
                    If UpdateRecord Then
                        Call Main.executeSQL("default", "update cclibraryFiles set FileTypeID=" & FileTypeID & ",filesize=" & fileSize & " where ID=" & ResourceRecordID)
                    End If
                    '
                    ImageSrc = kmaEncodeURL(ImageSrc)
                    FormDetails = FormDetails & GetFormRow_Files(fileSize, IconLink, IconOnClick, RecordName, ImageSrc, ModifiedDate, RowCount, EditLink, Description, FileExtension, RecordName, ImageSrc, ImageAlt, ImageWidthText, ImageHeightText, ResourceRecordID, currentFolderID, AllowThumbnails, FileTypeFilter, ThumbNailSrc, SourceMode, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn)
                    Main.NextCSRecord (cS)
                    RowCount = RowCount + 1
                Next
            End If
            Call Main.closecs(cS)
            '
            ' ----- If nothing found, print no files found
            '
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
            FormDetails = FormDetails & GetFormRow_Blank("", "", "", "", "", 0, RowCount, "", "", "BLANK", "", "", "", "", "", 0, currentFolderID, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn, ColumnCnt)
        Next
        '
        ' Upload link
        '
        If AllowLocalFileAdd Then
            '
            ' Upload Form
            '
            FormDetails = FormDetails & GetFormRow_Options(currentFolderID, topFolderPath, ColumnCnt, IsContentManagerFiles, IsContentManagerFolders, currentFolderHasModifyAccess)
            RowCount = RowCount + 1
        End If
        '
        ' Bottom border
        '
        FormDetails = FormDetails & "<tr class=""border""><td class=""border"" Colspan=" & (ColumnCnt) & ">" & Main.GetPanel("&nbsp;") & "</td></tr>"
        FormDetails = FormDetails & "</table>"
        '
        ' Create the FormFolders
        '
        FormFolders = GetRLNav(currentFolderID, topFolderPath, topFolderID)
        FormFolders = "<div class=""nav"">" & FormFolders & "</div>"
        'FormFolders = Main.GetPanelRev(FormFolders)
        '
        ' Assemble the form
        '
        hint = "900"
        GetForm = GetForm & "<table border=0 cellpadding=0 cellspacing=0 width=""100%""><tr>"
        If Not blockFolderNavigation Then
            GetForm = GetForm & "<td class=""nav ccPanel3DInput"">" & FormFolders & "<BR><img src=/ResourceLibrary/spacer.gif width=140 height=1></td>"
            GetForm = GetForm & "<td class=""navBorder ccPanel3D""><img src=/ResourceLibrary/spacer.gif width=5 height=1></td>"
        End If
        GetForm = GetForm & "<td class=""content"">" & FormDetails & "</td>"
        GetForm = GetForm & "</tr></Table>"
        GetForm = GetForm & ButtonBar
        GetForm = GetForm & Main.GetFormInputHidden("RowCount", RowCount)
        GetForm = GetForm & Main.GetUploadFormEnd
    End If
    '
    GetForm = "" _
        & vbCrLf & vbTab & "<div class=""ccLibrary"">" _
        & kmaIndent(GetForm) _
        & vbCrLf & vbTab & "</div>" _
        & ""
    '
    ' Help Link
    '
    GetForm = Main.GetHelpLink(42, "Using the Resource Library", "The Resource Library is a convenient place to store reusable content, such as images and downloads. Objects in the Library can be placed on any page. The Library itself can be added to any page on your site.") & GetForm
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetForm, hint=" & hint)
End Function
'
'=================================================================================
' Returns the Resource Library Row HTML.
'=================================================================================
'
Private Function GetFormRow_Folders(ignore0 As String, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Long, EditLink As String, Description As String, FileType As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Long, FolderID As Long, AllowEditColumn As Boolean, AllowPlaceColumn As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim RowClass As String
    Dim AnchorTag As String
    Dim ImageTag As String
    Dim CellStart As String
    Dim CellStart2 As String
    Dim CellStart5 As String
    Dim CellEnd As String
    Dim Icon As String
    Dim IconLinkStart As String
    Dim IconLinkEnd As String
    Dim DateString As String
    Dim InnerCell As String
    Dim CellStartRight As String
    '
    If (RowCount Mod 2) = 0 Then
        RowClass = "ccPanelRowOdd"
    Else
        RowClass = "ccPanelRowEven"
        End If
    '
    
    CellStart = "<td class=""left ccAdminSmall " & RowClass & """>"
    CellStartRight = "<td class=""right ccAdminSmall " & RowClass & """>"
    CellStart2 = "<td class=""left ccAdminSmall " & RowClass & """>"
    CellStart5 = "<td class=""left ccAdminSmall " & RowClass & """>"
    CellEnd = "</td>"
    '
    If ModifiedDate = CDate(0) Then
        DateString = "&nbsp;"
    Else
        DateString = FormatDateTime(ModifiedDate, vbShortDate)
        End If
    '
    GetFormRow_Folders = GetFormRow_Folders & "<tr class=""row " & RowClass & """>"
    GetFormRow_Folders = GetFormRow_Folders & CellStart & "&nbsp;" & CellEnd
    If AllowEditColumn Then
        GetFormRow_Folders = GetFormRow_Folders & CellStart & "&nbsp;" & CellEnd
    End If
    If AllowPlaceColumn Then
        GetFormRow_Folders = GetFormRow_Folders & CellStart & "&nbsp;" & CellEnd
    'Else
    '    GetFormRow_Folders = GetFormRow_Folders & CellStart & "&nbsp;" & CellEnd
    End If
    GetFormRow_Folders = GetFormRow_Folders & CellStart & "<A href=""" & Main.ServerPage & "?" & IconLink & """>" & IconFolderOpen & "</A>" & CellEnd
    GetFormRow_Folders = GetFormRow_Folders & CellStart & Name & CellEnd
    GetFormRow_Folders = GetFormRow_Folders & CellStart & Description & CellEnd
    GetFormRow_Folders = GetFormRow_Folders & CellStart & "&nbsp;" & CellEnd
    GetFormRow_Folders = GetFormRow_Folders & CellStartRight & DateString & CellEnd
    GetFormRow_Folders = GetFormRow_Folders & "</tr>"
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetFormRow_Folders")
End Function
'
'=================================================================================
' Returns the Resource Library Row HTML.
'=================================================================================
'
Private Function GetFormRow_ChildFolders(ignore0 As String, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Long, EditLink As String, Description As String, FileType As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Long, FolderID As Long, AllowEditColumn As Boolean, AllowPlaceColumn As Boolean, AllowSelectColumn As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim RowClass As String
    Dim AnchorTag As String
    Dim ImageTag As String
    Dim CellStart As String
    'Dim CellStart2 As String
    'Dim CellStart5 As String
    Dim CellEnd As String
    Dim Icon As String
    Dim IconLinkStart As String
    Dim IconLinkEnd As String
    Dim DateString As String
    Dim InnerCell As String
    Dim CellStartCenter As String
    Dim CellStartRight As String
    '
    If (RowCount Mod 2) = 0 Then
        RowClass = "ccPanelRowOdd"
    Else
        RowClass = "ccPanelRowEven"
    End If
    '
    CellStart = vbCrLf & "<td class=""left ccAdminSmall"">"
    CellStartCenter = vbCrLf & "<td class=""center ccAdminSmall"">"
    CellStartRight = vbCrLf & "<td class=""right ccAdminSmall"">"
    CellEnd = "</td>"
    '
    If ModifiedDate = CDate(0) Then
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
        GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartCenter & "<input type=checkbox ID=Select" & RowCount & " name=Row" & RowCount & " value=1 onClick=""RLRowClick(this.checked,'Row" & RowCount & "');"">" & Main.GetFormInputHidden("Row" & RowCount & "FolderID", FolderID) & CellEnd
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
    GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartCenter & "<A href=""" & Main.ServerPage & "?" & IconLink & """>" & IconFolderClosed & "</A>" & CellEnd
    GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStart & "<A href=""" & Main.ServerPage & "?" & IconLink & """>" & Name & "</A>" & CellEnd
    GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStart & Description & CellEnd
    GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartRight & "&nbsp;" & CellEnd
    GetFormRow_ChildFolders = GetFormRow_ChildFolders & CellStartRight & DateString & CellEnd
    GetFormRow_ChildFolders = GetFormRow_ChildFolders & "</tr>"
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetFormRow_ChildFolders")
End Function
'
'=================================================================================
' Returns the Resource Library Row HTML.
'=================================================================================
'
Private Function GetFormRow_Files(fileSize As Long, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Long, EditLink As String, Description As String, FilenameExt As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Long, FolderID As Long, AllowThumbnails As Boolean, FileTypeFilter As String, ThumbNailSrc As String, SourceMode As Long, AllowEditColumn As Boolean, AllowPlaceColumn As Boolean, AllowSelectColumn As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim ImageLink As String
    Dim JSCopy As String
    Dim RowClass As String
    Dim AnchorTag As String
    Dim ImageTag As String
    Dim CellStart As String
    Dim CellStartRight As String
    Dim CellStart2 As String
    Dim CellStart5 As String
    Dim CellEnd As String
    Dim IconIMG As String
    Dim IconLinkStart As String
    Dim IconLinkEnd As String
    Dim DateString As String
    Dim InnerCell As String
    Dim PreviewImageURL As String
    Dim CellStartCenter As String
    Dim FileTypePtr As Long
    Dim IconFilename As String
    Dim IsImage As Boolean
    Dim IsVideo As Boolean
    Dim IsFlash As Boolean
    Dim IsMedia As Boolean
    Dim Mediafilename As String
    Dim IsDownload As Boolean
    Dim Downloadfilename As String
    Dim FileTypeName As String
    Dim TestFileTYpe As String
    Dim FileTypeFound As Boolean
    Dim MediaIMG As String
    Dim JSClose As String
    '
    If (RowCount Mod 2) = 0 Then
        RowClass = "ccPanelRowOdd"
    Else
        RowClass = "ccPanelRowEven"
    End If
    '
    CellStart = vbCrLf & "<td class=""left ccAdminSmall"">"
    CellStartCenter = vbCrLf & "<td class=""center ccAdminSmall"">"
    CellStartRight = vbCrLf & "<td class=""right ccAdminSmall"">"
    CellEnd = "</td>"
    '
    If ModifiedDate = CDate(0) Then
        DateString = "&nbsp;"
    Else
        DateString = FormatDateTime(ModifiedDate, vbShortDate)
    End If
    '
    ' Determine Icons and actions
    '
Dim AllowPlace As Boolean
    AllowPlace = False
    If IconFileCnt <= 0 Then
        IconIMG = IconImage
    Else
        TestFileTYpe = "," & UCase(Replace(FilenameExt, ".", "")) & ","
        For FileTypePtr = 0 To IconFileCnt - 1
            
            If InStr(1, "," & IconFiles(FileTypePtr).ExtensionList & ",", UCase(TestFileTYpe), vbTextCompare) <> 0 Then
                    With IconFiles(FileTypePtr)
                        FileTypeName = .Name
                        IsImage = .IsImage
                        IsVideo = .IsVideo
                        IsFlash = .IsFlash
                        IsMedia = IsImage Or IsVideo Or IsFlash
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
                Exit For
            End If
        Next
    End If
    '
    If Not FileTypeFound Then
        FileTypeName = TestFileTYpe
        IsImage = False
        IsVideo = False
        IsFlash = False
        Mediafilename = ""
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
        GetFormRow_Files = GetFormRow_Files & CellStartCenter & "<input type=checkbox ID=Select" & RowCount & " name=Row" & RowCount & " value=1 onClick=""RLRowClick(this.checked,'Row" & RowCount & "');"">" & Main.GetFormInputHidden("Row" & RowCount & "FileID", RecordID) & CellEnd
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
        ' Allow selection of resources to be placed on the opening pages
        '
        If SelectLinkObjectName <> "" Then
            '
            ' return the objects URL to the input element with ID=SelectLinkObjectName
            '
            JSCopy = kmaEncodeJavascript(ResourceLink)
            ImageLink = "<img src=""/ResourceLibrary/ResourceLink1616.gif"" border=""0"" width=""16"" height=""16"" alt=""Place a link to this resource"" title=""Place a link to this resource"" valign=""absmiddle"">"
            GetFormRow_Files = GetFormRow_Files & CellStartCenter & "<a href=""#"" onClick=""var e=window.opener.document.getElementById('" & SelectLinkObjectName & "');e.value='" & JSCopy & "'; window.close();"">" & ImageLink & "</A>" & CellEnd
        ElseIf SourceMode = SourceModeFromDownloadRequest Then
            '
            ' return a simple download
            '
            If IsDownload Then
                JSCopy = Downloadfilename
                JSCopy = Replace(JSCopy, "\", "\\")
                JSCopy = kmaEncodeJavascript(JSCopy)
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
                JSCopy = kmaEncodeJavascript(ResourceLink)
                JSClose = "" _
                    & " if(navigator.appName.indexOf('Microsoft')!=-1) {window.returnValue='" & JSCopy & "'}" _
                    & " else{window.opener.setAssetValue('" & JSCopy & "')}" _
                    & " self.close();"
                GetFormRow_Files = GetFormRow_Files & CellStartCenter & "<a href=""#"" onClick=""" & JSClose & """ >" & MediaIMG & "</A>" & CellEnd
            Else
                GetFormRow_Files = GetFormRow_Files & CellStartCenter & IconNoFile & CellEnd
            End If
        End If
    End If
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
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetFormRow_Files")
End Function
''
''=================================================================================
'' Returns the Resource Library Row HTML.
''=================================================================================
''
'Private Function GetFormRow_AddChild(ignore0 As String, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Long, EditLink As String, Description As String, FileType As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Long, FolderID As Long) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim RowClass As String
'    Dim AnchorTag As String
'    Dim ImageTag As String
'    Dim CellStartCenter As String
'    Dim CellStart As String
'    Dim CellStart2 As String
'    Dim CellStart5 As String
'    Dim CellEnd As String
'    Dim Icon As String
'    Dim IconLinkStart As String
'    Dim IconLinkEnd As String
'    Dim DateString As String
'    Dim InnerCell As String
'    '
'    If IsContentManager Then
'        CellStart = "<td Style=""border-top: 1px solid #C0C0C0; Padding: 2px 2px 2px 2px"" align=""left"" valign=""middle""><SPAN class=""ccAdminSmall"">"
'        CellStartCenter = "<td Style=""border-top: 1px solid #C0C0C0; Padding: 2px 2px 2px 2px"" Class=""" & RowClass & """ align=""center"" valign=""middle""><SPAN class=""ccAdminSmall"">"
'        CellEnd = "</SPAN></td>"
'        '
'        If ModifiedDate = CDate(0) Then
'            DateString = "&nbsp;"
'        Else
'            DateString = FormatDateTime(ModifiedDate, vbShortDate)
'            End If
'        If EditLink <> "" Then
'            GetFormRow_AddChild = GetFormRow_AddChild & "<tr Style=""Padding: 2px 2px 2px 2px"" align=""left"">"
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & "&nbsp;" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & "&nbsp;" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStart & "<b>Add Folder</b>" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & "<A href=""" & EditLink & """ target=""_blank"">" & IconFolderAdd & "</A>" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & IconNoFile & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & IconNoFile & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & "&nbsp;" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & "&nbsp;" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & "</tr>"
'        End If
'    End If
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleRLTrapError("GetFormRow_AddChild")
'    End Function
''
''=================================================================================
'' Returns the Resource Library Row HTML.
''=================================================================================
''
'Private Function GetFormRow_AddFile(ignore0 As String, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Long, EditLink As String, Description As String, FileType As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Long, FolderID As Long) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim RowClass As String
'    Dim AnchorTag As String
'    Dim ImageTag As String
'    Dim CellStart As String
'    Dim CellStartCenter As String
'    Dim CellStart2 As String
'    Dim CellStart5 As String
'    Dim CellEnd As String
'    Dim Icon As String
'    Dim IconLinkStart As String
'    Dim IconLinkEnd As String
'    Dim DateString As String
'    Dim InnerCell As String
'    '
'    If IsContentManager Then
'        CellStart = "<td Style=""border-top: 1px solid #C0C0C0; Padding: 2px 2px 2px 2px"" align=""left"" valign=""middle""><SPAN class=""ccAdminSmall"">"
'        CellStartCenter = "<td Style=""border-top: 1px solid #C0C0C0; Padding: 2px 2px 2px 2px"" Class=""" & RowClass & """ align=""center"" valign=""middle""><SPAN class=""ccAdminSmall"">"
'        CellEnd = "</SPAN></td>"
'        '
'        If ModifiedDate = CDate(0) Then
'            DateString = "&nbsp;"
'        Else
'            DateString = FormatDateTime(ModifiedDate, vbShortDate)
'        End If
'        If EditLink <> "" Then
'            GetFormRow_AddFile = GetFormRow_AddFile & "<tr Style=""Padding: 2px 2px 2px 2px"" align=""left"">"
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "<b>Add File</b>" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStartCenter & "<A href=""" & EditLink & """ target=""_blank"">" & IconFileAdd & "</A>" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStartCenter & IconNoFile & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStartCenter & IconNoFile & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & "</tr>"
'        End If
'    End If
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleRLTrapError("GetFormRow_AddFile")
'    End Function
'
'=================================================================================
' Returns the Resource Library Row HTML.
'=================================================================================
'
Private Function GetFormRow_Blank(ignore0 As String, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Long, EditLink As String, Description As String, FileType As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Long, FolderID As Long, AllowEditColumn As Boolean, AllowPlaceColumn As Boolean, AllowSelectColumn As Boolean, ColumnCnt As Long) As String
    On Error GoTo ErrorTrap
    '
    GetFormRow_Blank = vbCrLf & vbTab & "<tr class=""listRow""><td class=""left""><img height=""23"" width=""22"" src=""/ResourceLibrary/spacer.gif""></td><td class=""left"" colspan=""" & ColumnCnt - 1 & """>&nbsp;</td></tr>"
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetFormRow_Blank")
End Function
'
'=================================================================================
' Returns the Resource Library Row HTML.
'=================================================================================
'
Private Function GetFormRow_Options(FolderID As Long, topFolderPath As String, ColumnCnt As Long, IsContentManagerFiles As Boolean, IsContentManagerFolders As Boolean, hasModifyAccess As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim moveSelect As String
    Dim RowClass As String
    Dim AnchorTag As String
    Dim ImageTag As String
    Dim Icon As String
    Dim FileCell As String
    Dim FolderCell As String
    Dim folderPtr As Long
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
            & "<table id=""AddFolderTable"" border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%"">" _
            & "<tr>"
        FolderCell = FolderCell _
            & "<td class=""left"" align=""left"" colspan=2>" & kmaAddSpan("Add Folder&nbsp;", "ccAdminSmall") & "<BR><img src=""/ResourceLibrary/spacer.gif"" width=""230"" height=""1""></td>" _
            & "<td class=""left"" Width=""99%"" align=""left"">" & kmaAddSpan("Description&nbsp;", "ccAdminSmall") & "<BR><img src=""/ResourceLibrary/spacer.gif"" width=""100"" height=""1""></td>" _
            & "</tr><tr>" _
            & "<td class=""left"" Width=""30"" align=""right"">1&nbsp;<BR><img src=/ResourceLibrary/spacer.gif width=30 height=1></td>" _
            & "<td class=""left"" align=""left""><INPUT TYPE=""Text"" NAME=""FolderName.1"" SIZE=""30""></td>" _
            & "<td class=""left"" align=""left""><INPUT TYPE=""Text"" NAME=""FolderDescription.1"" SIZE=""40""></td>" _
            & "</tr>"
        FolderCell = FolderCell _
            & "</Table>" _
            & "<table border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%"">" _
            & "<tr><td class=""left"" Width=""30""><img src=/ResourceLibrary/spacer.gif width=30 height=1></td><td align=""left""><a href=""#"" onClick=""InsertFolderRow(); return false;"">+ Add more folders</a></td></tr>" _
            & "</Table>" & Main.GetFormInputHidden("AddFolderCount", 1, "AddFolderCount")
    End If
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
            & "</Table>" & Main.GetFormInputHidden("LibraryUploadCount", 1, "LibraryUploadCount")
    End If
    '
    '
    '
    GetFormRow_Options = "" _
        & "<img src=""/ResourceLibrary/spacer.gif"" width=""1"" height=""5"">" _
        & "<BR>" & Main.GetFormInputCheckBox("AllowThumbnails", Main.GetMemberProperty("LibraryAllowthumbnails", False)) & "&nbsp;Display Thumbnails"
    If Main.IsAdmin Or hasModifyAccess Then
        moveSelect = GetMoveFolderPathSelect(FolderID, topFolderPath)
        If moveSelect <> "" Then
            GetFormRow_Options = GetFormRow_Options & "<BR>" & Main.GetFormInputCheckBox("Move", False, "Move") & "&nbsp;Move selected files to " & moveSelect
        End If
        If FolderCell <> "" Then
            GetFormRow_Options = GetFormRow_Options & "<BR><BR>" & Main.GetPanelInput(FolderCell)
        End If
        If FileCell <> "" Then
            GetFormRow_Options = GetFormRow_Options & "<BR>" & Main.GetPanelInput(FileCell)
        End If
    End If
    If GetFormRow_Options <> "" Then
        GetFormRow_Options = Main.GetPanel(GetFormRow_Options)
        GetFormRow_Options = "<tr><td class=""left"" colspan=" & (ColumnCnt) & ">" & GetFormRow_Options & "</td></tr>"
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetFormRow_Options")
End Function
'
'
'
Private Function GetForm_HeaderCell(Align As String, Width As String, Copy As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Style
    '
    Style = "" _
        & "padding: 3px;" _
        & "font-size:10px;"
    GetForm_HeaderCell = "<td WIDTH=""" & Width & """ ALIGN=""" & Align & """ class=ccAdminListCaption style=""" & Style & """>" _
        & Copy _
        & "</td>"
        '& Main.GetPanel(kmaAddSpan(Copy, "ccAdminSmall") & "<BR /><IMG SRC=""/ResourceLibrary/spacer.gif"" WIDTH=""1"" HEIGHT=""1"">")
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetForm_HeaderCell")
End Function
'
'
'
Private Function IsInFolder(topFolderID As Long, FolderID As Long, Optional ParentPath As String) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim cS As Long
    Dim ParentID As Long
    '
    If FolderID = 0 Then
        IsInFolder = False
    ElseIf topFolderID = 0 Then
        IsInFolder = True
    ElseIf (InStr(1, "," & ParentPath & ",", "," & CStr(FolderID) & ",") <> 0) Then
        IsInFolder = False
    Else
        ParentPath = ParentPath & "," & CStr(FolderID)
        cS = Main.OpenCSContentRecord("Library Folders", FolderID, , , "ParentID")
        If Main.IsCSOK(cS) Then
            ParentID = Main.GetCSInteger(cS, "ParentID")
            End If
        Call Main.closecs(cS)
        If ParentID = 0 Then
            IsInFolder = False
        ElseIf ParentID = topFolderID Then
            IsInFolder = True
        Else
            IsInFolder = IsInFolder(topFolderID, ParentID, ParentPath)
            End If
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetResourceLibraryRow")
End Function
'
'
'
Private Function GetParentFoldersLink(topFolderPath As String, topFolderID As Long, currentFolderID As Long, FolderID As Long, RefreshQS As String, ChildIDList As String) As String
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cS As Long
    Dim ParentID As Long
    Dim folderName As String
    Dim RecordFound As Boolean
    Dim FolderLink As String
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
        cS = Main.OpenCSContent("Library Folders", "ID=" & FolderID, , , , , "Name,ParentID")
        If Main.IsCSOK(cS) Then
            RecordFound = True
            ParentID = Main.GetCSInteger(cS, "ParentID")
            folderName = Main.GetCSText(cS, "Name")
        End If
        Call Main.closecs(cS)
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
            GetParentFoldersLink = GetParentFoldersLink(topFolderPath, topFolderID, currentFolderID, ParentID, RefreshQS, ChildIDList & "," & FolderID) & "\" & FolderLink
            'GetParentFoldersLink = GetParentFoldersLink(topFolderPath, topFolderID, CurrentFolderID, ParentID, RefreshQS, ChildIDList & "," & FolderID) & " > " & FolderLink
        End If
    End If
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetParentFoldersLink")
End Function
'
'----------------------------------------------------------------------------------------
'   Get a select menu of all folders with which you have ModifyAccess
'----------------------------------------------------------------------------------------
'
Private Function GetFolderPathSelect(topFolderPathID As Long, topFolderPath As String, RequireModifyAccess As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim optionCnt As Long
    Dim SQL As String
    Dim cS As Long
    Dim Ptr As Long
    Dim PtrFolderID As Long
    Dim cnt As Long
    Dim BakeName As String
    Dim Pos As Long
    Dim FullPath As String
    Dim IndexCnt As Long
    Dim parentFolderID As Long
    Dim FolderID As Long
    Dim PtrString As String
    Dim hasAccess As Boolean
    Dim pathRemoveString As String
    Dim pathCaption As String
    '
    'GetFolderPathSelect = FolderSelect
    If GetFolderPathSelect = "" Then
        '
        ' create full paths, set .hasViewAccess
        '
        optionCnt = 0
        If topFolderPath <> "" Then
            pathRemoveString = "root\"
            Pos = InStrRev(topFolderPath, "\")
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
        Ptr = FolderPathIndex.GetFirstPointer
        Do While (Ptr >= 0)
            If folders(Ptr).hasViewAccess And ((Not RequireModifyAccess) Or folders(Ptr).hasModifyAccess) Then
                PtrFolderID = folders(Ptr).FolderID
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
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("getFolderPathSelect")
End Function
'
'
'
Private Function GetFolderPath(targetPtr As Long, ChildIDList As String) As String
    On Error GoTo ErrorTrap
    '
    Dim ParentPtr As Long
    Dim ParentID As Long
    Dim FolderID As Long
    
    '
    GetFolderPath = folders(targetPtr).Name
    ParentID = folders(targetPtr).parentFolderID
    FolderID = folders(targetPtr).FolderID
    If ParentID = 0 Then
        '
        ' At the Root page
        '
        GetFolderPath = "Root\" & GetFolderPath
    ElseIf (FolderID = ParentID) Or (InStr(1, "," & ChildIDList & ",", "," & ParentID & ",") <> 0) Then
        '
        ' circular reference - Make this a root page b
        '
    Else
        For ParentPtr = 0 To UBound(folders)
            If folders(ParentPtr).FolderID = ParentID Then
                GetFolderPath = GetFolderPath(ParentPtr, ChildIDList & "," & ParentID) & "\" & GetFolderPath
                'GetFolderPath = GetFolderPath(ParentPtr, ChildIDList & "," & ParentID) & " > " & GetFolderPath
                Exit For
            End If
        Next
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("getFolderPath")
End Function
'
'
'
Private Function GetJumpFolderPathSelect(FolderID As Long, topFolderPath As String) As String
    On Error GoTo ErrorTrap
    '
    GetJumpFolderPathSelect = GetFolderPathSelect(FolderID, topFolderPath, False)
    If GetJumpFolderPathSelect <> "" Then
        GetJumpFolderPathSelect = Replace(GetJumpFolderPathSelect, "FieldName", "JumpFolderID")
        GetJumpFolderPathSelect = Replace(GetJumpFolderPathSelect, "onChange", "onChange=""QJump(this);"" ")
        GetJumpFolderPathSelect = Replace(GetJumpFolderPathSelect, "value=", "value=?" & Main.RefreshQueryString & "&FolderID=")
        GetJumpFolderPathSelect = "<script language=JavaScript1.2>function QJump(e){var l=e.value;if(l!=''){window.name='RL';window.location.assign(l);}}</script>" & GetJumpFolderPathSelect
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetJumpFolderPathSelect")
End Function

'
'
'
Private Function GetMoveFolderPathSelect(FolderID As Long, topFolderPath As String) As String
    On Error GoTo ErrorTrap
    '
    GetMoveFolderPathSelect = GetFolderPathSelect(FolderID, topFolderPath, True)
    GetMoveFolderPathSelect = Replace(GetMoveFolderPathSelect, "FieldName", "MoveFolderID")
    GetMoveFolderPathSelect = Replace(GetMoveFolderPathSelect, "onChange", "onChange=""var e=getElementById('Move');if(e){e.checked=true};"" ")
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetMoveFolderPathSelect")
End Function
'
'=============================================================
'
'=============================================================
'
Private Function GetRLNav(currentFolderID As Long, topFolderPath As String, topFolderID As Long) As String
    On Error GoTo ErrorTrap
    '
    '
    Dim cS As Long
    Dim SQL As String
    Dim Tree As New MenuTreeClass
    Dim BakeName As String
    Dim Caption As String
    Dim Id As Long
    Dim ParentID As Long
    Dim Link As String
    Dim IsAuthoring As Boolean
    Dim LinkBase As String
    Dim CurrentFolderName As String
    Dim Ptr As Long
    '
    IsAuthoring = False
    BakeName = "RLNav"
    If Not IsAuthoring Then
'        GetRLNav = Main.ReadBake(BakeName)
    End If
    If GetRLNav = "" Then
        LinkBase = Main.RefreshQueryString
        LinkBase = ModifyQueryString(LinkBase, "FolderID", "0")
        If topFolderID = 0 Then
            Call Tree.AddEntry(CStr(0), CStr(-1), , , "?" & LinkBase, "Root")
        End If
        If folderCnt > 0 Then
            For Ptr = 0 To folderCnt - 1
                Id = folders(Ptr).FolderID
                If folders(Ptr).hasViewAccess Then
                'If hasModifyAccessByFolder(Id, topFolderPath) Then
                    ParentID = folders(Ptr).parentFolderID
                    Caption = Replace(folders(Ptr).Name, " ", "&nbsp;")
                    Link = "?" & ModifyQueryString(LinkBase, "FolderID", CStr(Id))
                    Call Tree.AddEntry(CStr(Id), CStr(ParentID), , , Link, Caption)
                End If
            Next
        End If
        GetRLNav = Tree.GetTree(CStr(topFolderID), CStr(currentFolderID))
        Call Main.SaveBake(BakeName, GetRLNav, "Library Folders")
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
    Call Main.AddOnLoadJavascript("convertTrees(); expandToItem('tree0','" & currentFolderID & "');")
    'Link = "?" & LinkBase
    'Link = "<div style=""position:relative;left:-10;margin-bottom:3px;""><a href=" & Link & " style=""text-decoration:none ! important;"">" & topFolderPath & "</a></div>"
    'GetRLNav = Replace(GetRLNav, "<LI ", Link & "<LI ", 1, 1, vbTextCompare)
    ''If CurrentFolderID <> 0 Then
        'GetRLNav = GetRLNav & "<script type=""text/javascript"">convertTrees(); expandToItem('tree0','" & CurrentFolderID & "');</script>"
    ''End If
    
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetRLNav")
End Function
'
'
'
Private Function AllowFolderAccess(FolderID As Long, ParentID As Long)
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cS As Long
    Dim GrandParentID As Long
    '
    If FolderID = 0 Or Main.IsAdmin Then
        AllowFolderAccess = True
    Else
        SQL = "select top 1 *" _
            & " from ccMemberRules M,ccLibraryFolderRules R" _
            & " where M.MemberID=" & Main.memberID _
            & " and R.FolderID=" & FolderID _
            & " and M.GroupID=R.GroupID" _
            & " and R.Active<>0" _
            & " and M.Active<>0" _
            & " and ((M.DateExpires is null)or(M.DateExpires>" & KmaEncodeSQLDate(Now) & "))"
        cS = Main.OpenCSSQL("default", SQL)
        AllowFolderAccess = Main.IsCSOK(cS)
        Call Main.closecs(cS)
        '
        ' If no folder access, check its parent folder
        '
        If Not AllowFolderAccess And (ParentID <> 0) Then
            cS = Main.OpenCSContentRecord("Library Folders", ParentID)
            If Main.IsCSOK(cS) Then
                GrandParentID = Main.GetCSInteger(cS, "ParentID")
            End If
            Call Main.closecs(cS)
            AllowFolderAccess = AllowFolderAccess(ParentID, GrandParentID)
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("AllowFolderAccess")
End Function
'
'
'
Private Function hasModifyAccessByFolder(FolderID As Long, topFolderPath As String) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cS As Long
    Dim GrandParentID As Long
    Dim FolderCells As Variant
    Dim Ptr As Long
    '
    If FolderID = 86 Then
    FolderID = FolderID
    End If
    
    '
    If Main.IsAdmin Then
        '
        '
        '
        hasModifyAccessByFolder = True
    Else
        '
        ' Need to check permissions
        '
        Call LoadFolders_returnTopFolderId(topFolderPath)
        If FolderID = 0 Then
            hasModifyAccessByFolder = True
        Else
            Ptr = FolderIdIndex.GetPointer(CStr(FolderID))
            If Ptr >= 0 Then
                hasModifyAccessByFolder = folders(Ptr).hasModifyAccess
            End If
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("hasModifyAccessByFolder")
End Function
'
'
'
Private Function LoadFolders_returnTopFolderId(topFolderPath As String) As Long
    On Error GoTo ErrorTrap
    '
    Dim testFolderID As Long
    Dim testFullPath As String
    Dim testFolderPtr As Long
    Dim testParentID As Long
    Dim targetFolderId As Long
    Dim targetParentFolderID As Long
    Dim targetFolderName As String
    Dim targetFolders() As String
    Dim targetFolderCnt As Long
    Dim topFolderID As Long
    Dim lcasetopFolderPath As String
    Dim Ptr As Long
    Dim SQL As String
    Dim cS As Long
    Dim GrandParentID As Long
    Dim FolderCells As Variant
    Dim FolderID As Long
    Dim hasModifyAccess As Boolean
    Dim IsAdmin As Boolean
    '
    If (folderCnt = 0) Then
        IsAdmin = Main.IsAdmin
        lcasetopFolderPath = LCase(topFolderPath)
        '
        ' Load the folders storage
        '
        SQL = "select Distinct" _
            & " F.ID" _
            & " ,F.ParentID" _
            & " ,F.Name" _
            & " ,(select top 1 ID from ccMemberRules where ccMemberRules.MemberID=" & Main.memberID & " and ccMemberRules.GroupID=FR.GroupID) as Allowed" _
            & " from (cclibraryfolders F left join ccLibraryFolderRules FR on FR.FolderID=F.ID)" _
            & " where (f.active<>0)" _
            & " order by f.name"
        cS = Main.OpenCSSQL("default", SQL)
        If Main.IsCSOK(cS) Then
            FolderCells = Main.GetCSRows(cS)
        End If
        Call Main.closecs(cS)
        If Not IsEmpty(FolderCells) Then
            folderCnt = UBound(FolderCells, 2) + 1
        End If
        If folderCnt > 0 Then
            '
            ' Store folders and setup folder index
            '
            ReDim folders(folderCnt - 1)
            Set FolderIdIndex = New FastIndex5Class
            Set FolderNameIndex = New FastIndex5Class
            For Ptr = 0 To folderCnt - 1
                FolderID = kmaEncodeInteger(FolderCells(0, Ptr))
                Call FolderIdIndex.SetPointer(CStr(FolderID), Ptr)
                targetFolderName = kmaEncodeText(FolderCells(2, Ptr))
                Call FolderNameIndex.SetPointer(targetFolderName, Ptr)
                With folders(Ptr)
                    .FolderID = FolderID
                    .parentFolderID = kmaEncodeInteger(FolderCells(1, Ptr))
                    hasModifyAccess = IsAdmin Or (Not IsNull(FolderCells(3, Ptr)))
                    .Name = targetFolderName
                    .hasModifyAccess = hasModifyAccess
                    .modifyAccessIsValid = hasModifyAccess
                    .hasViewAccess = False
                End With
            Next
            '
            ' set FullPath, propigate modifyAccess from parent to folder , set ViewAccess
            '
            Set FolderPathIndex = New FastIndex5Class
            For Ptr = 0 To folderCnt - 1
                With folders(Ptr)
                    '
                    ' determine modify access
                    '
                    If (Not .modifyAccessIsValid) Then
                        .hasModifyAccess = LoadFolders_GetModifyAccess(.parentFolderID)
                        .modifyAccessIsValid = True
                    End If
                    '
                    ' set FullPath
                    '
                    'testFolderID = folders(Ptr).FolderID
                    'testFullPath = folders(Ptr).FullPath
                    'If testFullPath = "" Then
                    testFullPath = GetFolderPath(Ptr, "")
                    folders(Ptr).FullPath = testFullPath
                    'End If
                    Call FolderPathIndex.SetPointer(testFullPath, Ptr)
                    '
                    ' hasViewAccess
                    '
'                    If topFolderPath <> "" Then
'                        '
'                        ' block paths that are not within the topFolderPath
'                        '
'                        Pos = InStr(1, testFullPath, "\" & topFolderPath, vbTextCompare)
'                        If Pos = 0 Then
'                            testFullPath = ""
'                        Else
'                            testFullPath = Mid(testFullPath, Pos + 1)
'                        End If
'                    End If
                    If InStr(1, testFullPath, "root\" & topFolderPath, vbTextCompare) = 1 Then
                    'If LCase(testFullPath) = LCase("root\" & topFolderPath) Then
                        '
                        ' if this path is under the topFolderpath, viewAccess=true
                        '
                        .hasViewAccess = True
                    End If
                End With
            Next
            '
            ' determine topFolderID from topFolderPath
            ' go through targetfolder string from top down
            '
            topFolderID = 0
            If topFolderPath <> "" Then
                targetFolders = Split(topFolderPath, "\")
                targetFolderCnt = UBound(targetFolders) + 1
                topFolderID = loadFolders_getFolderID(targetFolders, targetFolderCnt - 1)
                '
                ' if topFolderId not found, create the new folder(s) necessary to targetFolderPath
                '
                If topFolderID = 0 Then
                    targetFolderId = 0
                    For Ptr = 0 To targetFolderCnt - 1
                        targetFolderName = targetFolders(Ptr)
                        targetParentFolderID = targetFolderId
                        '
                        ' find or create the folder with this name and this targetParentFolderID
                        '
                        testFolderPtr = FolderNameIndex.GetPointer(targetFolderName)
                        Do While testFolderPtr >= 0
                            testParentID = folders(testFolderPtr).parentFolderID
                            If targetParentFolderID <> testParentID Then
                                '
                                ' right name but wrong parent, try next
                                '
                            Else
                                '
                                ' good match, set this as the parent and find the next
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
                            cS = Main.InsertCSRecord("Library Folders")
                            If Main.IsCSOK(cS) Then
                                targetFolderId = Main.GetCSInteger(cS, "ID")
                                Call Main.SetCS(cS, "name", targetFolderName)
                                Call Main.SetCS(cS, "parentid", targetParentFolderID)
                            End If
                            Call Main.closecs(cS)
                        End If
                        If Ptr = (targetFolderCnt - 1) Then
                            topFolderID = targetFolderId
                        End If
                    Next
                End If
            End If
        End If
    End If
    '
    LoadFolders_returnTopFolderId = topFolderID
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("hasModifyAccessByFolder")
End Function
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
Private Function loadFolders_getFolderID(targetArray() As String, targetArrayPtr As Long) As Long
    On Error GoTo ErrorTrap
    '
    Dim cachePtr As Long
    Dim cacheFolderID As Long
    Dim cacheParentFolderID As Long
    Dim targetFolderName As String
    Dim targetFolderParentId As Long
    '
    loadFolders_getFolderID = 0
    targetFolderName = targetArray(targetArrayPtr)
    cachePtr = FolderNameIndex.GetPointer(targetFolderName)
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
            targetFolderParentId = loadFolders_getFolderID(targetArray, targetArrayPtr - 1)
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
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("loadFolders_getFolderID")
End Function
'
'
'
Private Function LoadFolders_GetModifyAccess(FolderID As Long) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim Ptr As Long
    '
    Ptr = FolderIdIndex.GetPointer(CStr(FolderID))
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
            LoadFolders_GetModifyAccess = LoadFolders_GetModifyAccess(folders(Ptr).parentFolderID)
            folders(Ptr).hasModifyAccess = LoadFolders_GetModifyAccess
            folders(Ptr).modifyAccessIsValid = True
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("LoadFolders_GetModifyAccess")
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
    Set Main = Nothing
    End Sub
'
'
'
Private Function GetFileSize(VirtualFilePathPage As String) As Long
    On Error GoTo ErrorTrap
    '
    Dim FileDescriptor As String
    Dim FileSplit As String
    Dim FileSplit2() As String
    Dim FileParts() As String
    Dim Ptr As Long
    Dim SlashPosition As Long
    Dim Filename As String
    Dim Pathname As String
    Dim tickstart As Long
    Dim hint As String
    '
hint = "1"
    tickstart = GetTickCount
    'Call AppendLogFile("GetFileSize, VirtualFilePathPage=" & VirtualFilePathPage)
    '
hint = "2"
    VirtualFilePathPage = Replace(VirtualFilePathPage, "/", "\")
    SlashPosition = InStrRev(VirtualFilePathPage, "\")
    If SlashPosition <> 0 Then
        Filename = LCase(Mid(VirtualFilePathPage, SlashPosition + 1))
        Pathname = Mid(VirtualFilePathPage, 1, SlashPosition - 1)
    End If
    FileDescriptor = Main.GetFileList(Pathname)
hint = "3"
    If FileDescriptor = "" Then
        'Call AppendLogFile("GetFileSize, descriptor is blank")
    Else
hint = "4"
        FileSplit2 = Split(FileDescriptor, vbCrLf)
        'Call AppendLogFile("GetFileSize, FileDescriptor lines=" & UBound(FileSplit2))
hint = "5"
        For Ptr = 0 To UBound(FileSplit2)
            FileParts = Split(FileSplit2(Ptr), ",")
            If UBound(FileParts) <= 5 Then
                'Call AppendLogFile("GetFileSize, FileDescriptor row [" & Ptr * "] has <6 parts, descrriptor=" & FileDescriptor)
            Else
                If LCase(FileParts(0)) = Filename Then
                    GetFileSize = kmaEncodeInteger(FileParts(5))
                    'Call AppendLogFile("GetFileSize, match on " & FileParts(0))
                    Exit For
                End If
            End If
        Next
hint = "6"
    End If
    'Call AppendLogFile("GetFileSize, exit, filesize=" & GetFileSize)
    '
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetFileSize, hint=" & hint)
End Function
'
'
'
Private Function GetFileTypeID(Filename As String) As Long
    On Error GoTo ErrorTrap
    '
    Dim FileNameSplit() As String
    Dim FileExtension As String
    Dim CSType As Long
    Dim DefaultFileTypeID As Long
    Dim cnt As Long
    Dim Ptr As Long
    Dim hint As String
    '
    FileNameSplit = Split(Filename, ".")
    FileExtension = FileNameSplit(UBound(FileNameSplit))
    '
    ' try to read if from IconFiles
    '
    hint = "1"
    cnt = IconFileCnt
    If cnt > 0 Then
        For Ptr = 0 To cnt - 1
            hint = "2"
            If InStr(1, "," & IconFiles(Ptr).ExtensionList & ",", "," & FileExtension & ",", vbTextCompare) <> 0 Then
                hint = "3"
                GetFileTypeID = IconFiles(Ptr).FileTypeID
                Exit For
                If LCase(IconFiles(Ptr).Name) = "default" Then
                    hint = "4"
                    DefaultFileTypeID = IconFiles(Ptr).FileTypeID
                End If
            End If
        Next
        If Ptr = cnt Then
            GetFileTypeID = DefaultFileTypeID
        End If
    End If
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
    Exit Function
ErrorTrap:
    Call HandleRLTrapError("GetFileTypeID, hint=" & hint)
End Function
'
'
'
Private Sub HandleRLTrapError(MethodName As String, Optional ignore0 As String)
    Set Main = Nothing
    Call HandleError("ResourceLibrary", MethodName, Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'
'
Private Sub HandleRLInternalError(MethodName As String, ErrDescription As String)
    Call Err.Raise(KmaErrorInternal, App.EXEName, ErrDescription)
    End Sub

