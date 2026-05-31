using System;
using System.Collections.Generic;
using System.Xml;
using Contensive.BaseClasses;
using Contensive.Models.Db;
using Contensive.Addons.ResourceLibrary.Controllers;
using Contensive.Addons.ResourceLibrary.Models.Db;
using Contensive.Addons.ResourceLibrary.Models.Domain;
using static Contensive.Addons.ResourceLibrary.constants;
using static Contensive.Addons.ResourceLibrary.Controllers.genericController;

namespace Contensive.Addons.ResourceLibrary.Views {
    //
    public class LegacyLibraryClass {
        //
        public List<FileTypeModel> iconFiles = new List<FileTypeModel>();
        //
        public FolderTypeModel[] folders = Array.Empty<FolderTypeModel>();
        public int folderCnt;
        public Controllers.FastIndexClass FolderIdIndex = new Controllers.FastIndexClass();
        public Controllers.FastIndexClass FolderNameIndex = new Controllers.FastIndexClass();
        public Controllers.FastIndexClass FolderPathIndex = new Controllers.FastIndexClass();
        //
        public string FolderSelect;
        //
        // -----------------------------------------------------------------------------------
        // ----- Publics
        // -----------------------------------------------------------------------------------
        // ----- not used
        //
        public int UserMemberID;
        public string RequestStream;
        //
        // ----- Icons used
        //
        public const string IconFolderOpen = "<img src=\"/ResourceLibrary/IconFolderOpen.gif\" border=\"0\" width=\"22\" height=\"23\" ALT=\"Close this folder\">";
        public const string IconFolderClosed = "<img src=\"/ResourceLibrary/IconFolderClosed.gif\" border=\"0\" width=\"22\" height=\"23\" alt=\"Open this folder\">";
        public const string IconFolderAdd = "<img src=\"/ResourceLibrary/IconFolderAdd2.gif\" border=\"0\" width=\"22\" height=\"23\" alt=\"Add a New folder\">";
        public const string IconFolderEdit = "<img src=\"/ResourceLibrary/IconFolderEdit.gif\" border=\"0\" width=\"22\" height=\"23\" alt=\"Edit this folder\">";
        public const string IconFile = "<img src=\"/ResourceLibrary/IconFile.gif\" border=\"0\" width=\"22\" height=\"23\" alt=\"file\">";
        public const string IconFileAdd = "<img src=\"/ResourceLibrary/IconContentAdd.gif\" border=\"0\" width=\"18\" height=\"22\" alt=\"Add a New  file\">";
        public const string IconFileEdit = "<img src=\"/ResourceLibrary/IconContentEdit.gif\" border=\"0\" width=\"18\" height=\"22\" alt=\"Edit this file\">";
        public const string IconPreview = "<img src=\"/ResourceLibrary/IconPreview.gif\" border=\"0\" width=\"22\" height=\"23\" alt=\"Preview this image\">";
        public const string IconDownload = "<img src=\"/ResourceLibrary/IconDownload3.gif\" border=\"0\" width=\"16\" height=\"16\" alt=\"Select this download\" valign=\"absmiddle\">";
        public const string IconCreateImage = "<img src=\"/ResourceLibrary/IconimagePlace.gif\" border=\"0\" width=\"18\" height=\"22\" alt=\"Select this image\">";
        public const string IconCreateDownload = "<img src=\"/ResourceLibrary/IconDownload3.gif\" border=\"0\" width=\"16\" height=\"16\" alt=\"Select this download\" valign=\"absmiddle\">";
        public const string IconSpacer = "<img src=\"/ResourceLibrary/spacer.gif\" width=\"22\" height=\"23\">";
        public const string IconView = "<img src=\"/ResourceLibrary/IconView.gif\" border=\"0\" width=\"22\" height=\"23\" alt=\"Preview this file\">";
        public const string IconImage = "<img src=\"/ResourceLibrary/IconImage2.gif\" border=\"0\" width=\"22\" height=\"23\" alt=\"Image\">";
        public const string IconPDF = "<img src=\"/ResourceLibrary/IconPDF.gif\" border=\"0\" width=\"16\" height=\"16\" alt=\"Adobe Pdf\">";
        public const string IconOther = "<img src=\"/ResourceLibrary/IconFile.gif\" border=\"0\" width=\"22\" height=\"23\" alt=\"Unrecognized File Type\">";
        public const string IconNoFile = "<img src=/ResourceLibrary/BulletRound2.gif width=5 height=5>";
        //
        // ----- SelectResource Support
        //       This means the resource library supports buttons that allow objects to be
        //       placed on different page from the resource library, like an Editor
        //
        //
        // ----- If an editor is used to call the resource library, the window.opener.insertresource()
        //       call needs the object name of the editor so the contents can be copied to the invisible
        //       form field after the changes (no onchange event available)
        //
        public string SelectResourceEditorObjectName;
        //
        // ----- If AllowPlace is true and SelectLinkObjectName<>"", the RL is being used as a link selector
        //       When the 'place' icon is clicked, the URL of the resource is copied to the window.opener.[selectlinkobjectname]
        //
        public string SelectLinkObjectName;
        //
        // ----- Blocks the folder list in the left hand side
        //
        public bool blockFolderNavigation;
        //
        // -----------------------------------------------------------------------------------
        // ----- Privates
        // -----------------------------------------------------------------------------------
        //
        public int iMinRows;
        public int iFolderID;                      // Current Folder being Displayed, 0 for root
        public int SourceMode;                      //
        //
        //        // SourceMode
        //        //   3/6/2010 - moved codes up to capture the 0 case and it to page
        //        //   1 = From Editor Object or Link selector: allow image and download insert, provide close button
        //        //   2 = From Editor Image Properties: allow image insert, provide close button
        //        //   3 = From Admin site, no inserts, and provide cancel button
        //const SourceModeOnPage = 1
        //const SourceModeFromDownloadRequest = 2
        //const SourceModeFromLinkDialog = 3
        //   0 = From Editor Object selector: allow image and download insert, provide close button
        //   1 = From Editor Image Properties: allow image insert, provide close button
        //   2 = From Admin site, no inserts, and provide cancel button
        public const int SourceModeFromDownloadRequest = 0;
        public const int SourceModeFromLinkDialog = 1;
        public const int SourceModeOnPage = 2;
        //
        //   0 caller is the editor directly, clicking on action icons calls InsertImaage, etc
        //   1 caller is the editor image page, clicking on action icons calls the image page methods
        //
        public int HoldPosition;
        //Private main As MainClass

        //
        //=====================================================================================
        //// <summary>
        //// AddonDescription
        //// </summary>
        //// <param name="CP"></param>
        //// <returns></returns>
        //Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
        //    Dim result As String = ""
        //    'Dim sw As New Stopwatch : sw.Start()
        //    Try
        //        '
        //        ' -- initialize application. If authentication needed and not login page, pass true
        //        Using ae As New applicationController(CP, False)
        //            main = New vbConversion.Contensive.VbConversion.MainClass(CP)
        //            '
        //            ' -- your code
        //            result = GetContent(CP)
        //            If ae.packageErrorList.Count > 0 Then
        //                result = "Hey user, this happened - " & Join(ae.packageErrorList.ToArray, "<br>")
        //            End If
        //        End Using
        //    Catch ex As Exception
        //        CP.Site.ErrorReport(ex)
        //    End Try
        //    Return result
        //End Function

        //
        //=================================================================================
        //   Aggregate Object Interface
        //=================================================================================
        //
        public string getResourceLibrary(CPBaseClass cp) {
            try {
                using (applicationController ae = new applicationController(cp)) {
                    //
                    SelectResourceEditorObjectName = cp.Doc.GetText("SelectResourceEditorObjectName");
                    SelectLinkObjectName = cp.Doc.GetText("SelectLinkObjectName");
                    blockFolderNavigation = cp.Utils.EncodeBoolean(cp.Doc.GetBoolean("Block Folder Navigation"));
                    //
                    return GetForm(cp, ae);
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return "There was an error displaying the resource library";
            }
        }

        //
        //=================================================================================
        // Returns the Resource Library HTML.
        //   This HTML does not include the HTML, HEAD or BODY tags.
        //=================================================================================
        //
        private string GetForm(CPBaseClass cp, applicationController ae) {
            string result = "";
            try {
                const string LibraryFileTypespathFilename = "resourcelibrary\\LibraryConfig.xml";
                string hint = "000";
                //
                string ButtonBarStyle = ""
                        + " color: black;"
                        + " font-weight: bold;"
                        + " padding: 5px;"
                        + " background-color: #a0a0a0;"
                        + " border-bottom: 1px solid #e0e0e0;"
                        + " border-right: 1px solid #e0e0e0;"
                        + " border-top: 1px solid #808080;"
                        + " border-left: 1px solid #808080;";
                //
                string OptionPanelStyle = ""
                        + " color: black;"
                        + " font-weight: bold;"
                        + " padding: 5px;"
                        + " background-color: #d0d0d0;"
                        + " border-bottom: 1px solid #e0e0e0;"
                        + " border-right: 1px solid #e0e0e0;"
                        + " border-top: 1px solid #a0a0a0;"
                        + " border-left: 1px solid #a0a0a0;";
                //
                if (!(false)) {
                    //
                    // Determine Current Folder
                    //
                    hint = "001";
                    //BuildVersion = cp.Site.GetText("build version")
                    bool IsContentManagerFiles = cp.User.IsContentManager("Library Files");
                    bool IsContentManagerFolders = cp.User.IsContentManager("Library Folders");
                    string Button = cp.Doc.GetText("Button");
                    string FileTypeFilter = cp.Doc.GetText("ffilter").ToLower();
                    cp.Doc.AddRefreshQueryString("ffilter", FileTypeFilter);
                    bool AllowThumbnails = cp.User.GetBoolean("LibraryAllowthumbnails", false);
                    string FolderIDString = cp.Doc.GetText("folderid");
                    int currentFolderID = cp.Utils.EncodeInteger(FolderIDString);
                    if (FolderIDString != "") {
                        cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString());
                    } else {
                        currentFolderID = cp.User.GetInteger("Libraryfolderid", "0");
                    }
                    //
                    // Load Folder cache
                    //
                    hint = $"010, topFolderPath={ae.topFolderPath}";
                    int topFolderID = LoadFolders_returnTopFolderId(cp, ae.topFolderPath);
                    //
                    bool reloadFolderCache = false;
                    int currentFolderPtr;
                    //
                    // verify that current folder has viewAccess (if not jumpt to root)
                    //
                    if (currentFolderID != 0) {
                        currentFolderPtr = FolderIdIndex.getPtr(currentFolderID.ToString());
                        if ((currentFolderPtr > (folders.Length - 1)) || (currentFolderPtr < 0)) {
                            currentFolderPtr = 0;
                        }
                        if (currentFolderID < 0) {
                            currentFolderID = 0;
                            cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString());
                        } else if (!folders[currentFolderPtr].hasViewAccess) {
                            currentFolderID = 0;
                            cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString());
                        }
                    }
                    //
                    // determine if current folder has modify access
                    //
                    hint = "020";
                    bool currentFolderHasModifyAccess = false;
                    if ((cp.User.IsAdmin || IsContentManagerFiles || IsContentManagerFolders)) {
                        //
                        // you get modify access if you can modify the content
                        //
                        currentFolderHasModifyAccess = true;
                    } else if (currentFolderID == 0) {
                        //
                        // only admin and content managers of files and folders have modify access to root folder
                        //
                    } else {
                        //
                        // others have modify access to this folder if they are in a modify access group
                        //
                        currentFolderPtr = FolderIdIndex.getPtr(currentFolderID.ToString());
                        if (currentFolderPtr >= 0) {
                            currentFolderHasModifyAccess = folders[currentFolderPtr].hasModifyAccess;
                        }
                    }
                    //topFolderID = GetFolderID(topFolderPath)
                    //
                    // Load IconFiles
                    //
                    hint = "030";
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(cp.File.ReadVirtual(LibraryFileTypespathFilename));
                    int Ptr;
                    hint = "040";
                    if ((doc.DocumentElement.Name.ToLower().Equals("libraryconfig"))) {
                        if (doc.DocumentElement.ChildNodes.Count > 0) {
                            Ptr = 0;
                            hint = "050";
                            foreach (XmlElement baseNode in doc.DocumentElement.ChildNodes) {
                                hint = "060";
                                switch (baseNode.Name.ToLower()) {
                                    case "filetype":
                                        hint = "070";
                                        Ptr = Ptr + 1;
                                        FileTypeModel newFileType = new FileTypeModel();
                                        iconFiles.Add(newFileType);
                                        //int IconCnt;
                                        //If Ptr >= IconCnt Then
                                        //    IconCnt = IconCnt + 10
                                        //    ReDim Preserve IconFiles(IconCnt)
                                        //End If
                                        foreach (XmlElement typeNode in baseNode.ChildNodes) {
                                            switch (typeNode.Name.ToLower()) {
                                                case "name":
                                                    newFileType.Name = typeNode.Value;
                                                    break;
                                                case "filetypeid":
                                                    newFileType.FileTypeID = cp.Utils.EncodeInteger(typeNode.Value);
                                                    break;
                                                case "extensionlist":
                                                    newFileType.ExtensionList = typeNode.Value;
                                                    break;
                                                case "isdownload":
                                                    newFileType.IsDownload = cp.Utils.EncodeBoolean(typeNode.Value);
                                                    break;
                                                case "isimage":
                                                    newFileType.IsImage = cp.Utils.EncodeBoolean(typeNode.Value);
                                                    break;
                                                case "isvideo":
                                                    newFileType.IsVideo = cp.Utils.EncodeBoolean(typeNode.Value);
                                                    break;
                                                case "isflash":
                                                    newFileType.IsFlash = cp.Utils.EncodeBoolean(typeNode.Value);
                                                    break;
                                                case "iconlink":
                                                    newFileType.IconFilename = typeNode.Value;
                                                    break;
                                                case "mediaiconlink":
                                                    newFileType.MediaIconFilename = typeNode.Value;
                                                    break;
                                                case "downloadiconlink":
                                                    newFileType.DownloadIconFilename = typeNode.Value;
                                                    break;
                                            }
                                        }
                                        break;
                                }
                            }
                        }
                    }
                    //
                    // Verify default icons
                    //
                    hint = "100";
                    string DefaultIcon = "\\cclib\\images\\IconImage2.gif";
                    string DefaultMedia = "\\cclib\\images\\Iconimage2Media.gif";
                    string DefaultDownload = "\\cclib\\images\\Iconimage2Download.gif";
                    //
                    if (cp.Doc.GetText("SourceMode") == "") {
                        SourceMode = SourceModeOnPage;
                    } else {
                        SourceMode = cp.Doc.GetInteger("SourceMode");
                    }
                    cp.Doc.AddRefreshQueryString("SourceMode", SourceMode.ToString());
                    //
                    // ----- verify currentFolderID
                    //
                    if (currentFolderID == 0) {
                        //
                        // No folder give, use root folder, no owner
                        currentFolderID = topFolderID;
                        cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString());
                    }
                    Models.Db.LibraryFolderModel folder = DbBaseModel.create<Models.Db.LibraryFolderModel>(cp, currentFolderID);
                    //string FolderGroupName;
                    int FolderParentID = 0;
                    if ((folder != null)) {
                        FolderParentID = folder.parentId;
                    }
                    if ((topFolderID != currentFolderID) && (topFolderID != FolderParentID)) {
                        //
                        // Check if Folder is under the given root folder
                        if (!IsInFolder(cp, topFolderID, currentFolderID)) {
                            //
                            // Current folder is not in Root Folder, Use Root Folder
                            currentFolderID = topFolderID;
                            cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString());
                        }
                    }
                    //
                    // ----- Determine if select, place or edit icons are allowed
                    //
                    int ColumnCnt = 5;
                    bool AllowPlaceColumn = ae.allowPlace && ((SourceMode == SourceModeFromLinkDialog) || (SourceMode == SourceModeFromDownloadRequest));
                    if (AllowPlaceColumn) {
                        ColumnCnt = ColumnCnt + 1;
                    }
                    bool AllowEditColumn = (IsContentManagerFiles || IsContentManagerFolders);
                    if (AllowEditColumn) {
                        ColumnCnt = ColumnCnt + 1;
                    }
                    bool AllowSelectColumn = currentFolderHasModifyAccess;
                    if (AllowSelectColumn) {
                        ColumnCnt = ColumnCnt + 1;
                    }
                    //
                    // ----- Setup folder editing
                    bool AllowFolderAuthoring = IsContentManagerFolders;
                    int FolderCID = 0;
                    if (AllowFolderAuthoring) {
                        FolderCID = cp.Content.GetID("Library Folders");
                    }
                    //
                    // ----- Setup file editing
                    bool AllowFileAuthoring = IsContentManagerFiles;
                    int FileCID = 0;
                    if (AllowFileAuthoring) {
                        FileCID = cp.Content.GetID("Library Files");
                    }
                    //int FolderGroupID;
                    //
                    // ----- Setup Local File Management
                    //       Allow if Content Manager or user has group membership
                    //       Always allow, everyone has access to the root folder, then if you can get to the folder, let em upload
                    bool AllowLocalFileAdd = true;
                    //
                    // ----- Process input
                    int Pos;
                    string AltSizeList;
                    string FilenameNoExtension;
                    string FileExtension;
                    int FileTypeID;
                    int RowCount;
                    if (Button != "") {
                        AllowThumbnails = cp.Doc.GetBoolean("AllowThumbnails");
                        cp.User.SetProperty("LibraryAllowthumbnails", AllowThumbnails.ToString());
                        switch (Button) {
                            case ButtonCancel:
                                //
                                // CAncel button, just redirect back to the current page
                                cp.Response.Redirect("#");
                                break;
                            case ButtonDelete:
                                //
                                //
                                //
                                RowCount = cp.Doc.GetInteger("RowCount");
                                int DeleteFileID;
                                int DeleteFolderID;
                                if (RowCount > 0) {
                                    for (Ptr = 0; Ptr <= RowCount - 1; Ptr++) {
                                        if (cp.Doc.GetBoolean($"Row{Ptr}")) {
                                            DeleteFolderID = cp.Doc.GetInteger($"Row{Ptr}FolderID");
                                            if (DeleteFolderID != 0) {
                                                //Call Main.WriteStream("Deleting Folder " & FolderID)
                                                cp.Content.Delete("Library Folders", $"id={DeleteFolderID}");
                                                reloadFolderCache = true;
                                            }
                                            DeleteFileID = cp.Doc.GetInteger($"Row{Ptr}FileID");
                                            if (DeleteFileID != 0) {
                                                //Call Main.WriteStream("Deleting File " & FileID)
                                                cp.Content.Delete("Library Files", $"id={DeleteFileID}");
                                                reloadFolderCache = true;
                                            }
                                        }
                                    }
                                }
                                break;
                            case ButtonApply:
                                //
                                // Move Files
                                //
                                if (cp.Doc.GetBoolean("Move")) {
                                    int targetFolderId = cp.Doc.GetInteger("MoveFolderID");
                                    RowCount = cp.Doc.GetInteger("RowCount");
                                    if (RowCount > 0) {
                                        for (Ptr = 0; Ptr <= RowCount - 1; Ptr++) {
                                            if (cp.Doc.GetBoolean($"Row{Ptr}")) {
                                                int MoveFolderID = cp.Doc.GetInteger($"Row{Ptr}FolderID");
                                                int MoveFileID = cp.Doc.GetInteger($"Row{Ptr}FileID");
                                                if (MoveFolderID != 0) {
                                                    cp.Db.ExecuteSQL($"update ccLibraryFolders set ParentID={targetFolderId} where ID={MoveFolderID}");
                                                    reloadFolderCache = true;
                                                } else if (MoveFileID != 0) {
                                                    cp.Db.ExecuteSQL($"update ccLibraryFiles set FolderID={targetFolderId} where ID={MoveFileID}");
                                                    reloadFolderCache = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                //
                                // Upload
                                //
                                if (AllowLocalFileAdd) {
                                    //
                                    // Add Folders
                                    //
                                    hint = "300";
                                    int AddFolderCount = cp.Doc.GetInteger("AddFolderCount");
                                    int UploadPointer;
                                    for (UploadPointer = 1; UploadPointer <= AddFolderCount; UploadPointer++) {
                                        string folderName = cp.Doc.GetText($"FolderName.{UploadPointer}");
                                        if (folderName != "") {
                                            if (IsContentManagerFolders && (!cp.User.IsAdmin) && (currentFolderID == 0)) {
                                                //
                                                // Content Managers can not add folders to the root folder
                                                //
                                                cp.UserError.Add("Your account does not have access to add new folders to the root folder.");
                                                break;
                                            } else {
                                                Models.Db.LibraryFolderModel libraryFolder = DbBaseModel.addDefault<Models.Db.LibraryFolderModel>(cp);
                                                libraryFolder.name = folderName;
                                                libraryFolder.description = cp.Doc.GetText($"FolderDescription.{UploadPointer}");
                                                libraryFolder.parentId = currentFolderID;
                                                libraryFolder.save(cp);
                                                //cS = Main.InsertCSRecord("Library Folders")
                                                //If Main.IsCSOK(cS) Then
                                                //    Copy = cp.Doc.GetText("FolderDescription." & UploadPointer)
                                                //    Call Main.SetCS(cS, "Name", folderName)
                                                //    Call Main.SetCS(cS, "Description", Copy)
                                                //    If currentFolderID <> 0 Then
                                                //        Call Main.SetCS(cS, "ParentID", currentFolderID)
                                                //    End If
                                                //End If
                                                //Call Main.closecs(cS)
                                                reloadFolderCache = true;
                                            }
                                        }
                                    }
                                    //
                                    // Upload files
                                    //
                                    hint = "400";
                                    int UploadCount = cp.Doc.GetInteger("LibraryUploadCount");
                                    string uploadFilename = "";
                                    //int imagefileFolderId = cp.Doc.GetInteger("FolderID");
                                    for (UploadPointer = 1; UploadPointer <= UploadCount; UploadPointer++) {
                                        string imageRequestName = $"{RequestNameLibraryUpload}.{UploadPointer}";
                                        uploadFilename = cp.Doc.GetText($"{RequestNameLibraryUpload}.{UploadPointer}");
                                        if (uploadFilename != "") {
                                            hint = "410";
                                            LibraryFileModel libraryFile = DbBaseModel.addDefault<LibraryFileModel>(cp);


                                            string libraryName = cp.Doc.GetText($"{RequestNameLibraryName}.{UploadPointer}");
                                            if (libraryName == "") {
                                                libraryName = uploadFilename;
                                            }
                                            libraryFile.name = libraryName;
                                            string libraryDescription = cp.Doc.GetText($"{RequestNameLibraryDescription}.{UploadPointer}");
                                            if (libraryDescription == "") {
                                                libraryDescription = uploadFilename;
                                            }
                                            FileExtension = "";
                                            FilenameNoExtension = "";
                                            AltSizeList = "";
                                            Pos = (uploadFilename.LastIndexOf(".") + 1);
                                            if (Pos > 0) {
                                                FileExtension = uploadFilename.Substring(Pos - 1 + 1);
                                                FilenameNoExtension = uploadFilename.Substring(0, Pos - 1);
                                            }
                                            string cdnPathFilename = cp.Db.CreateUploadFieldPathFilename(LibraryFileModel.tableMetadata.tableNameLower, "filename", libraryFile.id, uploadFilename, CPContentBaseClass.FieldTypeIdEnum.File);
                                            string cdnPath = cdnPathFilename.Replace(uploadFilename, "");
                                            // string cdnPath = cp.CdnFiles.GetPath(cdnPathFilename);
                                            libraryFile.description = libraryDescription;
                                            libraryFile.folderId = currentFolderID;
                                            cp.Html.ProcessInputFile(imageRequestName, cdnPath);

                                            libraryFile.fileSize = GetFileSize(cp, cp.CdnFiles.PhysicalFilePath + libraryFile.name);
                                            string FileType = "";
                                            hint = "425";
                                            FileTypeID = GetFileTypeID(cp, uploadFilename);
                                            libraryFile.fileTypeId = FileTypeID;
                                            libraryFile.name = libraryName;
                                            libraryFile.description = libraryDescription;
                                            libraryFile.filename = cdnPath + uploadFilename;
                                            libraryFile.modifiedDate = DateTime.Now;
                                            libraryFile.save(cp);

                                            reloadFolderCache = true;
                                        }
                                    }
                                }
                                break;
                        }
                    }
                    hint = "500";
                    if (reloadFolderCache) {
                        folderCnt = 0;
                        topFolderID = LoadFolders_returnTopFolderId(cp, ae.topFolderPath);
                        reloadFolderCache = false;
                    }
                    //
                    // Housekeep potential issue where a parent is deleted and child does not show
                    //
                    string SQL = "update cclibraryfolders"
                            + " Set parentid=null"
                            + " where id in"
                            + " ("
                            + " select c.id from (cclibraryfolders c left join cclibraryfolders p on p.id=c.parentid)"
                            + " where p.ID Is Null"
                            + " and c.parentid is not null"
                            + " and c.parentid<>0"
                            + " )";
                    cp.Db.ExecuteSQL(SQL);
                    //
                    // Housekeep potential issue where a folder deleted and file does not show
                    //
                    SQL = "update cclibraryfiles"
                            + " Set folderid=null"
                            + " where id in"
                            + " ("
                            + " select c.id from (cclibraryfiles c left join cclibraryfolders p on p.id=c.folderid)"
                            + " where p.ID Is Null"
                            + " and c.folderid is not null"
                            + " and c.folderid<>0"
                            + " )";
                    cp.Db.ExecuteSQL(SQL);
                    //
                    // ----- Begin output
                    //
                    string rnbutton = "Button";
                    string ButtonExit;
                    if ((SourceMode == SourceModeFromDownloadRequest) || (SourceMode == SourceModeFromLinkDialog)) {
                        ButtonExit = cp.Html.Button(rnbutton, ButtonClose, "", "windowcloseID");
                    } else {
                        ButtonExit = cp.Html.Button(rnbutton, ButtonCancel);
                    }
                    string ButtonBar = "";
                    if (AllowLocalFileAdd) {
                        if (currentFolderHasModifyAccess) {
                            ButtonBar = "<div class=ccAdminButtonBar>"
                                    + ButtonExit
                                    + cp.Html.Button(rnbutton, ButtonApply)
                                    + cp.Html.Button(rnbutton, ButtonDelete, RequestNameButton, "returnDeleteCheckID")
                                    + "</div>";
                        } else {
                            ButtonBar = "<div class=ccAdminButtonBar>"
                                    + ButtonExit
                                    + cp.Html.Button(rnbutton, ButtonApply)
                                    + "</div>";
                        }
                    }

                    //result +=  genericController.htmlHidden("FolderID", currentFolderID)
                    result += ButtonBar;

                    string JumpSelect = "";
                    JumpSelect = GetJumpFolderPathSelect(cp, currentFolderID, ae.topFolderPath);
                    result += $"<div style=\"padding:10px;\">{GetParentFoldersLink(cp, ae.topFolderPath, topFolderID, currentFolderID, currentFolderID, cp.Doc.RefreshQueryString, "")}</div>";
                    if (JumpSelect != "") {
                        result += $"<div style=\"padding:10px;padding-top:0px\">Jump to&nbsp;{JumpSelect}</div>";
                    }
                    //
                    // From here down the form divides into FormFolder and FormDetails
                    //
                    string FormDetails = $"<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\"><tr class=\"headRow\">";
                    if (AllowSelectColumn) {
                        FormDetails += GetForm_HeaderCell(cp, "center", "10", $"Select<BR>{spacer1x10}");
                    }
                    if (AllowEditColumn) {
                        FormDetails += GetForm_HeaderCell(cp, "center", "15", $"Edit<br>{spacer1x15}");
                    }
                    if (AllowPlaceColumn) {
                        FormDetails += GetForm_HeaderCell(cp, "center", "15", $"Place<br>{spacer1x15}");
                    }
                    FormDetails = FormDetails
                            + GetForm_HeaderCell(cp, "left", "20", $"&nbsp;<BR>{spacer1x20}")
                            + GetForm_HeaderCell(cp, "left", "20%", $"Name<br>{spacer1x20}")
                            + GetForm_HeaderCell(cp, "left", "50%", $"Description<br>{spacer1x15}")
                            + GetForm_HeaderCell(cp, "center", "50", $"Size<br>{spacer1x50}")
                            + GetForm_HeaderCell(cp, "center", "50", $"Modified&nbsp;&nbsp;<br>{spacer1x50}")
                            + "</tr>";
                    //
                    // ----- Select the Folder Rows
                    //
                    string Criteria = "((ParentID is null)or(ParentID=0))";
                    //
                    if (currentFolderID != 0) {
                        cp.Doc.AddRefreshQueryString("FolderID", currentFolderID.ToString());
                    }
                    //
                    string SortField = cp.Doc.GetText("sortfield");
                    if (SortField == "") {
                        SortField = "Name";
                    }
                    cp.Doc.AddRefreshQueryString("SortField", SortField);
                    //
                    int SortDirection = cp.Doc.GetInteger("sortdirection");
                    if (SortDirection != 0) {
                        cp.Doc.AddRefreshQueryString("SortDirection", SortDirection.ToString());
                    }
                    //
                    if (SortDirection != 0 && SortField != "") {
                        SortField = SortField + " DESC";
                    }
                    //
                    Models.Db.LibraryFolderModel parentFolder = null;

                    int parentFolderID = 0;
                    if (currentFolderID != 0) {
                        //
                        // ----- FolderID given, lookup record and get ParentID
                        //       Note that allowupfolder allows users to "up" past top if they set it manually
                        //       Fix this when security is added
                        //
                        folder = DbBaseModel.create<Models.Db.LibraryFolderModel>(cp, currentFolderID);
                        if ((folder != null)) {
                            parentFolderID = folder.parentId;
                        }
                        parentFolder = DbBaseModel.create<Models.Db.LibraryFolderModel>(cp, parentFolderID);
                        Criteria = $"(ParentID={KmaEncodeSQLNumber(cp, currentFolderID)})";
                    } else if (ae.topFolderPath != "") {
                        //
                        // ----- Rootfolder given, lookup record and get ParentID
                        //
                        folder = DbBaseModel.createByUniqueName<Models.Db.LibraryFolderModel>(cp, ae.topFolderPath);
                        if ((folder != null)) {
                            parentFolderID = 0;
                            currentFolderID = folder.id;
                            cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString());
                        }
                        parentFolder = DbBaseModel.create<Models.Db.LibraryFolderModel>(cp, parentFolderID);
                        Criteria = $"(ParentID={KmaEncodeSQLNumber(cp, currentFolderID)})";
                    } else {
                        //
                        // ----- Use Root as top (no record)
                        parentFolder = DbBaseModel.create<Models.Db.LibraryFolderModel>(cp, parentFolderID);
                        //
                    }
                    //
                    // ----- Output the page
                    //
                    RowCount = 0;
                    hint = "700";
                    if (true) {
                        //
                        // ----- List out the folders
                        List<Models.Db.LibraryFolderModel> folderList = DbBaseModel.createList<Models.Db.LibraryFolderModel>(cp, Criteria, SortField);
                        string IconLink;
                        string EditLink;
                        DateTime? ModifiedDate;
                        foreach (var folder2 in folderList) {
                            string ChildFolderName = folder2.name;
                            if (ChildFolderName == "") {
                                ChildFolderName = "[no name]";
                            }
                            EditLink = "";
                            if (AllowFolderAuthoring) {
                                EditLink = $"{adminUrl(cp)}?cid={FolderCID}&id={folder2.id}&af=4&aa=2&depth=1";
                            }
                            IconLink = cp.Utils.ModifyQueryString(cp.Doc.RefreshQueryString, "folderid", folder2.id.ToString());
                            ModifiedDate = folder2.modifiedDate;
                            if (ModifiedDate <= DateTime.MinValue) {
                                ModifiedDate = folder2.dateAdded;
                            }
                            int ChildFolderID;
                            FormDetails += GetFormRow_ChildFolders(cp, IconFolderClosed, IconLink, "", ChildFolderName, "", ModifiedDate ?? DateTime.MinValue, RowCount, EditLink, folder2.description, "CHILD", "", "", "", "", "", 0, folder2.id, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn);
                            RowCount = RowCount + 1;
                        }

                        //
                        // Lookup the files in the folder
                        //
                        hint = "720";
                        if (currentFolderID == 0) {
                            Criteria = "((FolderID is null)or(FolderID=0))";
                        } else {
                            Criteria = $"(FolderID={KmaEncodeSQLNumber(cp, currentFolderID)})";
                        }
                        //FieldList = "ID,Name,ModifiedDate,Filename,Width,Height,DateAdded,Description,AltText,FileTypeID,FileSize,AltSizeList"
                        if (currentFolderID == 0) {
                            Criteria = "((FolderID is null)or(FolderID=0))";
                        } else {
                            Criteria = $"(FolderID={KmaEncodeSQLNumber(cp, currentFolderID)})";
                        }
                        List<LibraryFileModel> fileList = LibraryFileModel.createList<LibraryFileModel>(cp, Criteria, SortField);
                        foreach (var file in fileList) {
                            bool UpdateRecord = false;
                            int ResourceRecordID = file.id;
                            string RecordName = file.name;
                            ModifiedDate = file.modifiedDate;
                            string Filename = file.filename;
                            string ImageWidthText = file.width.ToString();
                            string ImageHeightText = file.height.ToString();
                            if (ModifiedDate <= DateTime.MinValue) {
                                ModifiedDate = file.dateAdded;
                            }
                            string Description = file.description;
                            string ImageAlt = file.altText;
                            FileTypeID = file.fileTypeId;
                            int fileSize = file.fileSize;
                            AltSizeList = file.altSizeList;
                            //
                            string ImageSrc = cp.Site.FilePath + Filename.Replace("\\", "/");
                            //
                            int DotPosition = (ImageSrc.LastIndexOf(".") + 1);
                            if (DotPosition == 0) {
                                FileExtension = "";
                                FilenameNoExtension = "";
                            } else {
                                FileExtension = ImageSrc.Substring(DotPosition - 1 + 1).ToUpper();
                                FilenameNoExtension = ImageSrc.Substring(1 - 1, DotPosition - 1);
                            }
                            //
                            if (FileTypeID == 0) {
                                FileTypeID = GetFileTypeID(cp, ImageSrc);
                                if (FileTypeID != 0) {
                                    UpdateRecord = true;
                                }
                            }
                            //
                            // if no name given, use the filename
                            //
                            if (RecordName == "") {
                                if (ImageSrc == "") {
                                    RecordName = "[no name]";
                                } else {
                                    DotPosition = (ImageSrc.LastIndexOf("/") + 1);
                                    if (DotPosition == 0) {
                                        RecordName = ImageSrc;
                                    } else {
                                        RecordName = ImageSrc.Substring(DotPosition - 1 + 1);
                                    }
                                }

                            }
                            file.name = RecordName;

                            //
                            string ResourceHref = "";
                            IconLink = "";
                            if (AllowFileAuthoring) {
                                EditLink = $"{adminUrl(cp)}?cid={FileCID}&id={ResourceRecordID}&af=4&aa=2&depth=1";
                            } else {
                                EditLink = "";
                            }
                            string ThumbNailSrc = "";
                            //
                            // create thumbnail
                            //
                            if (AllowThumbnails) {
                                ThumbNailSrc = ImageSrc;
                                if ((FilenameNoExtension != "") && (AltSizeList != "")) {
                                    string[] AltSizes = AltSizeList.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                                    int BestFitHeight = 9999999;
                                    string BestFitAltSize = "";
                                    for (Ptr = 0; Ptr <= (AltSizes.Length - 1); Ptr++) {
                                        //
                                        // Find the smallest image larger then height 50
                                        //
                                        string AltSize = AltSizes[Ptr].Trim();
                                        if (AltSize != "") {
                                            Pos = (AltSize.IndexOf("x", StringComparison.Ordinal) + 1);
                                            if (Pos > 0) {
                                                int AltSizeHeight = cp.Utils.EncodeInteger(AltSize.Substring(Pos - 1 + 1));
                                                if (AltSizeHeight >= 50 && AltSizeHeight < BestFitHeight) {
                                                    BestFitHeight = AltSizeHeight;
                                                    BestFitAltSize = AltSize;
                                                }
                                            }
                                        }
                                    }
                                    if (BestFitAltSize != "") {
                                        ThumbNailSrc = $"{FilenameNoExtension}-{BestFitAltSize}.{FileExtension}";
                                    }
                                    //
                                    //
                                    //
                                }
                            }
                            //
                            // -- get file size
                            if (fileSize == 0) {
                                string Pathname = cp.CdnFiles.PhysicalFilePath + Filename.Replace("/", "\\");
                                fileSize = GetFileSize(cp, Pathname);
                                if (fileSize != 0) {
                                    UpdateRecord = true;
                                }
                            }
                            //
                            //
                            //
                            if (UpdateRecord) {
                                cp.Db.ExecuteNonQuery($"update cclibraryFiles set FileTypeID={FileTypeID},filesize={fileSize} where ID={ResourceRecordID}");
                            }
                            //
                            ImageSrc = kmaEncodeURL(cp, ImageSrc);
                            string IconOnClick = "";
                            FormDetails += GetFormRow_Files(cp, fileSize, IconLink, IconOnClick, RecordName, ImageSrc, ModifiedDate ?? DateTime.MinValue, RowCount, EditLink, Description, FileExtension, RecordName, ImageSrc, ImageAlt, ImageWidthText, ImageHeightText, ResourceRecordID, currentFolderID, AllowThumbnails, FileTypeFilter, ThumbNailSrc, SourceMode, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn);
                            RowCount = RowCount + 1;
                        }
                        //
                        // ----- If nothing found, print no files found
                        if (RowCount == 0) {
                            FormDetails += $"<tr class=\"listRow\"><td class=\"center\">{IconSpacer}</td><td class=\"left\" colspan={ColumnCnt - 1}>no folders or files were found</td></tr>";
                            RowCount = RowCount + 1;
                        }
                    }
                    //
                    // Fill out the table to MinRows
                    //
                    hint = "800";
                    for (RowCount = RowCount; RowCount <= iMinRows; RowCount++) {
                        FormDetails += GetFormRow_Blank(cp, "", "", "", "", "", null, RowCount, "", "", "BLANK", "", "", "", "", "", 0, currentFolderID, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn, ColumnCnt);
                    }
                    //
                    // Upload link
                    //
                    if (AllowLocalFileAdd) {
                        //
                        // Upload Form
                        //
                        FormDetails += GetFormRow_Options(cp, currentFolderID, ae.topFolderPath, ColumnCnt, IsContentManagerFiles, IsContentManagerFolders, currentFolderHasModifyAccess);
                        RowCount = RowCount + 1;
                    }
                    //
                    // Bottom border
                    //
                    //FormDetails +=  "<tr class=\"border\"><td class=\"border\" Colspan=" & (ColumnCnt) & ">" & cp.Html.div("&nbsp;") & "</td></tr>"
                    FormDetails += "</table>";
                    //
                    // Create the FormFolders
                    //
                    string FormFolders = GetRLNav(cp, currentFolderID, ae.topFolderPath, topFolderID);
                    FormFolders = $"<div class=\"rlnav\">{FormFolders}</div>";
                    //FormFolders = Main.GetPanelRev(FormFolders)
                    //
                    // Assemble the form
                    //
                    hint = "900";
                    result += "<table border=0 cellpadding=0 cellspacing=0 width=\"100%\"><tr>";
                    if (!blockFolderNavigation) {
                        result += $"<td class=\"nav ccPanel3DInput\">{FormFolders}<BR><img src=/ResourceLibrary/spacer.gif width=140 height=1></td>";
                        result += "<td class=\"navBorder ccPanel3D\"><img src=/ResourceLibrary/spacer.gif width=5 height=1></td>";
                    }
                    result += $"<td class=\"content\">{FormDetails}</td>";
                    result += "</tr></Table>";
                    result += ButtonBar;
                    result += htmlHidden("RowCount", RowCount);
                    result = cp.Html.Form(result);
                }
                //
                // moved to layout -- result = "<div class=\"ccLibrary\">" & result & "</div>"
                //
                // Help Link
                //
                //result = Main.GetHelpLink(42, "Using the Resource Library", "The Resource Library is a convenient place to store reusable content, such as images and downloads. Objects in the Library can be placed on any page. The Library itself can be added to any page on your site.") & GetForm
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //=================================================================================
        // Returns the Resource Library Row HTML.
        //=================================================================================
        //
        private string GetFormRow_Folders(CPBaseClass cp, string ignore0, string IconLink, string IconOnClick, string Name, string NameLink, DateTime ModifiedDate, int RowCount, string EditLink, string Description, string FileType, string ResourceName, string ResourceLink, string ImageAlt, string ImageWidth, string ImageHeight, int RecordID, int FolderID, bool AllowEditColumn, bool AllowPlaceColumn) {
            string result = "";
            //
            try {
                string RowClass;
                //
                if ((RowCount % 2) == 0) {
                    RowClass = "ccPanelRowOdd";
                } else {
                    RowClass = "ccPanelRowEven";
                }
                //

                string CellStart = $"<td class=\"left  {RowClass}\">";
                string CellStartRight = $"<td class=\"right  {RowClass}\">";
                string CellStart2 = $"<td class=\"left  {RowClass}\">";
                string CellStart5 = $"<td class=\"left  {RowClass}\">";
                string CellEnd = "</td>";
                string DateString;
                //
                if (ModifiedDate <= DateTime.MinValue) {
                    DateString = "&nbsp;";
                } else {
                    DateString = ModifiedDate.ToShortDateString();
                }
                //
                result += $"<tr class=\"row {RowClass}\">";
                result += $"{CellStart}&nbsp;{CellEnd}";
                if (AllowEditColumn) {
                    result += $"{CellStart}&nbsp;{CellEnd}";
                }
                if (AllowPlaceColumn) {
                    result += $"{CellStart}&nbsp;{CellEnd}";
                    //Else
                    //    result +=  CellStart & "&nbsp;" & CellEnd
                }
                result += $"{CellStart}<A href=\"?{cp.Utils.EncodeUrl(IconLink)}\">{IconFolderOpen}</A>{CellEnd}";
                result += $"{CellStart}{Name}{CellEnd}";
                result += $"{CellStart}{Description}{CellEnd}";
                result += $"{CellStart}&nbsp;{CellEnd}";
                result += $"{CellStartRight}{DateString}{CellEnd}";
                result += "</tr>";
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //=================================================================================
        // Returns the Resource Library Row HTML.
        //=================================================================================
        //
        private string GetFormRow_ChildFolders(CPBaseClass cp, string ignore0, string IconLink, string IconOnClick, string Name, string NameLink, DateTime ModifiedDate, int RowCount, string EditLink, string Description, string FileType, string ResourceName, string ResourceLink, string ImageAlt, string ImageWidth, string ImageHeight, int RecordID, int FolderID, bool AllowEditColumn, bool AllowPlaceColumn, bool AllowSelectColumn) {
            string result = "";
            //
            try {
                //
                string RowClass;
                //
                if ((RowCount % 2) == 0) {
                    RowClass = "ccPanelRowOdd";
                } else {
                    RowClass = "ccPanelRowEven";
                }
                //
                string CellStart = $"\r\n<td class=\"left \">";
                string CellStartCenter = $"\r\n<td class=\"center \">";
                string CellStartRight = $"\r\n<td class=\"right \">";
                string CellEnd = "</td>";
                string DateString;
                //
                if (ModifiedDate <= DateTime.MinValue) {
                    DateString = "&nbsp;";
                } else {
                    DateString = ModifiedDate.ToShortDateString();
                }
                if (Description == "") {
                    Description = "&nbsp;";
                }
                //
                result = result + $"\r\n<tr class=\"listRow\" ID=\"Row{RowCount}\">";
                if (AllowSelectColumn) {
                    result = result + CellStartCenter + $"<input type=checkbox ID=Select{RowCount} name=Row{RowCount} value=1 onClick=\"RLRowClick(this.checked,'Row{RowCount}');\">" + htmlHidden($"Row{RowCount}FolderID", FolderID) + CellEnd;
                }
                if (AllowEditColumn) {
                    if (EditLink != "") {
                        result = result + CellStartCenter + $"<A href=\"{EditLink}\">{IconFolderEdit}</A>" + CellEnd;
                    } else {
                        result = result + CellStart + "&nbsp;" + CellEnd;
                    }
                }
                if (AllowPlaceColumn) {
                    result = result + CellStartCenter + IconNoFile + CellEnd;
                    //Else
                    //    result = result & CellStartCenter & IconNoFile & CellEnd
                }
                result = result + CellStartCenter + $"<A href=\"?{IconLink}\">{IconFolderClosed}</A>" + CellEnd;
                result = result + CellStart + $"<A href=\"?{IconLink}\">{Name}</A>" + CellEnd;
                result = result + CellStart + Description + CellEnd;
                result = result + CellStartRight + "&nbsp;" + CellEnd;
                result = result + CellStartRight + DateString + CellEnd;
                result = result + "</tr>";
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //=================================================================================
        // Returns the Resource Library Row HTML.
        //=================================================================================
        //
        private string GetFormRow_Files(CPBaseClass cp, int fileSize, string IconLink, string IconOnClick, string Name, string NameLink, DateTime ModifiedDate, int RowCount, string EditLink, string Description, string FilenameExt, string ResourceName, string ResourceLink, string ImageAlt, string ImageWidth, string ImageHeight, int RecordID, int FolderID, bool AllowThumbnails, string FileTypeFilter, string ThumbNailSrc, int SourceMode, bool AllowEditColumn, bool AllowPlaceColumn, bool AllowSelectColumn) {
            string result = "";
            //
            try {
                string RowClass;
                //
                if ((RowCount % 2) == 0) {
                    RowClass = "ccPanelRowOdd";
                } else {
                    RowClass = "ccPanelRowEven";
                }
                //
                string CellStart = $"\r\n<td class=\"left \">";
                string CellStartCenter = $"\r\n<td class=\"center \">";
                string CellStartRight = $"\r\n<td class=\"right \">";
                string CellStartRightPad = $"\r\n<td class=\"right  pr-3\">";
                string CellEnd = "</td>";
                string DateString;
                //
                if (ModifiedDate <= DateTime.MinValue) {
                    DateString = "&nbsp;";
                } else {
                    DateString = ModifiedDate.ToShortDateString();
                }
                //
                // Determine Icons and actions
                //
                bool AllowPlace;
                AllowPlace = false;
                string IconIMG = "";
                string IconFilename = "";
                bool IsImage = false;
                bool IsVideo = false;
                bool IsFlash = false;
                bool IsDownload = false;
                string Downloadfilename = "";
                string FileTypeName = "";
                string TestFileTYpe = "";
                bool FileTypeFound = false;
                string MediaIMG = "";
                if (iconFiles.Count <= 0) {
                    IconIMG = IconImage;
                } else {
                    TestFileTYpe = $",{FilenameExt.Replace(".", "").ToUpper()},";
                    foreach (FileTypeModel iconFile in iconFiles) {
                        int FileTypePtr = 0;
                        if ((iconFiles[FileTypePtr].ExtensionList.IndexOf(TestFileTYpe.ToUpper(), StringComparison.OrdinalIgnoreCase) + 1) != 0) {
                            FileTypeName = iconFile.Name;
                            IsImage = iconFile.IsImage;
                            IsVideo = iconFile.IsVideo;
                            IsFlash = iconFile.IsFlash;
                            bool IsMedia = IsImage || IsVideo || IsFlash;
                            //
                            // 4/15/08 - if no filter, show everything
                            //
                            //MediaIMG = IconNoFile

                            //                        If FileTypeFilter = "image" And IsImage Then
                            //                            MediaIMG = "<img src=""/ResourceLibrary/IconImagePlace2322.gif"" width=23 height=22 border=0 alt=""Place this image on the page"">"
                            //                            AllowPlace = True
                            //                        ElseIf FileTypeFilter = "media" And IsVideo Then
                            //                            MediaIMG = "<img src=""/ResourceLibrary/IconImagePlace2322.gif"" width=23 height=22 border=0 alt=""Place this video on the page"">"
                            //                            AllowPlace = True
                            //                        ElseIf FileTypeFilter = "flash" And IsVideo Then
                            //                            MediaIMG = "<img src=""/ResourceLibrary/IconImagePlace2322.gif"" width=23 height=22 border=0 alt=""Place this flash on the page"">"
                            //                            AllowPlace = True
                            //                        Else
                            //                            MediaIMG = IconNoFile
                            //                            AllowPlace = False
                            //                        End If
                            if (iconFile.MediaIconFilename != "") {
                                MediaIMG = $"<img src=\"{iconFile.MediaIconFilename}\" width=23 height=22 border=0 alt=\"Place this flash on the page\">";
                            }
                            IsDownload = iconFile.IsDownload;
                            Downloadfilename = iconFile.DownloadIconFilename;
                            IconFilename = iconFile.IconFilename;
                            if (IconFilename == "") {
                                IconFilename = "/ResourceLibrary/IconDefault.gif";
                            }
                            IconIMG = $"<img src=\"{IconFilename}\" border=\"0\" width=\"22\" height=\"23\" alt=\"{iconFile.Name}\">";
                            FileTypeFound = true;
                            break;
                        }
                    }
                    //For FileTypePtr = 0 To IconFileCnt - 1

                    //Next
                }
                //
                if (!FileTypeFound) {
                    if (FilenameExt == "PNG") {
                        IsImage = true;
                    } else if (FilenameExt == "JPG") {
                        IsImage = true;
                    } else if (FilenameExt == "GIF") {
                        IsImage = true;
                    } else {
                        IsImage = false;
                    }
                    FileTypeName = TestFileTYpe;
                    IsVideo = false;
                    IsFlash = false;
                    string Mediafilename = "";
                    IsDownload = true;
                    Downloadfilename = "/ResourceLibrary/IconDefaultDownload.gif";
                    IconFilename = "/ResourceLibrary/IconFile.gif";
                    IconIMG = IconOther;
                    MediaIMG = IconNoFile;
                }
                AllowPlace = false;
                if (FileTypeFilter == "image") {
                    if (IsImage) {
                        AllowPlace = true;
                    }
                } else if (FileTypeFilter == "media") {
                    if (IsVideo) {
                        AllowPlace = true;
                    }
                } else if (FileTypeFilter == "flash") {
                    if (IsFlash) {
                        AllowPlace = true;
                    }
                } else {
                    //
                    // no filter - place anything
                    //
                    AllowPlace = true;
                }
                if (AllowPlace && MediaIMG == "") {
                    MediaIMG = "<img src=\"/ResourceLibrary/IconImagePlace2322.gif\" width=23 height=22 border=0 alt=\"Place this file on the page\">";
                }
                //
                //   Output the row
                //
                result = result + $"\r\n<tr class=\"listRow\" ID=\"Row{RowCount}\">";
                if (AllowSelectColumn) {
                    result = result + CellStartCenter + $"<input type=checkbox ID=Select{RowCount} name=Row{RowCount} value=1 onClick=\"RLRowClick(this.checked,'Row{RowCount}');\">" + htmlHidden($"Row{RowCount}FileID", RecordID) + CellEnd;
                }
                //
                // ----- Edit Column
                //
                if (AllowEditColumn) {
                    if (EditLink != "") {
                        result = result + CellStartCenter + $"<A href=\"{EditLink}\">{IconFileEdit}</A>" + CellEnd;
                    } else {
                        result = result + CellStart + "&nbsp;" + CellEnd;
                    }
                }
                //
                // ----- Place Column
                //
                if (!AllowPlaceColumn) {
                    //
                    // hide column
                    //
                } else if ((!AllowPlace)) {
                    //
                    // Can not select resources - display dot
                    //
                    result = result + CellStartCenter + IconNoFile + CellEnd;
                } else {
                    //
                    string ImageLink;
                    string JSCopy;
                    //
                    // Allow selection of resources to be placed on the opening pages
                    //
                    if (SelectLinkObjectName != "") {
                        //
                        // return the objects URL to the input element with ID=SelectLinkObjectName
                        //
                        JSCopy = kmaEncodeJavascript(cp, ResourceLink);
                        ImageLink = "<img src=\"/ResourceLibrary/resourceLink1616.gif\" border=\"0\" width=\"16\" height=\"16\" alt=\"Place a link to this resource\" title=\"Place a link to this resource\" valign=\"absmiddle\">";
                        result = result + CellStartCenter + $"<a href=\"#\" onClick=\"var e=window.opener.document.getElementById('{SelectLinkObjectName}');e.value='{JSCopy}'; window.close();\">{ImageLink}</A>" + CellEnd;
                    } else if (SourceMode == SourceModeFromDownloadRequest) {
                        //
                        // return a simple download
                        //
                        if (IsDownload) {
                            JSCopy = Downloadfilename;
                            JSCopy = JSCopy.Replace("\\", "\\\\");
                            JSCopy = kmaEncodeJavascript(cp, JSCopy);
                            ImageLink = "<img src=\"/ResourceLibrary/IconDownload2.gif\" border=\"0\" width=\"23\" height=\"22\" alt=\"Link to this resource\" title=\"Link to this resource\" valign=\"absmiddle\">";
                            result = result + CellStartCenter + $"<a href=\"#\" onClick=\"window.opener.InsertDownload( '{RecordID}','{SelectResourceEditorObjectName}','{JSCopy}'); window.close();\">{ImageLink}</A>" + CellEnd;
                        } else {
                            result = result + CellStartCenter + IconNoFile + CellEnd;
                        }
                    } else if (SourceMode == SourceModeFromLinkDialog) {
                        //
                        // Return the file as a url to the editor dialog
                        //
                        if (AllowPlace) {
                            JSCopy = kmaEncodeJavascript(cp, ResourceLink);
                            string JSClose = ""
                            + $" if(navigator.appName.indexOf('Microsoft')!=-1) {{window.returnValue='{JSCopy}'}}"
                            + $" else{{window.opener.setAssetValue('{JSCopy}')}}"
                            + " self.close();";
                            result = result + CellStartCenter + $"<a href=\"#\" onClick=\"{JSClose}\" >{MediaIMG}</A>" + CellEnd;
                        } else {
                            result = result + CellStartCenter + IconNoFile + CellEnd;
                        }
                    }
                }
                NameLink = cp.Utils.DecodeUrl(NameLink);

                result = result + CellStartCenter + IconIMG + CellEnd;
                result = result + CellStart + $"<a href=\"{NameLink}\" target=\"_blank\">{Name}</A>" + CellEnd;
                //
                if (Description == "") {
                    Description = "&nbsp;";
                }
                if (AllowThumbnails && IsImage) {
                    //If AllowThumbnails And (UCase(FileTypeName) = "IMAGE") Then
                    result = result
                        + CellStart
                        + $"<a href=\"{NameLink}\" target=\"_blank\">"
                        + $"<img src=\"{ThumbNailSrc}\" height=\"50\"  vspace=\"0\" hspace=\"10\" style=\"vertical-align:middle;border:0;\">"
                        + "</a>"
                        + Description
                        + CellEnd;
                } else {
                    result = result
                        + CellStart
                        + Description
                        + CellEnd;
                }
                //
                if (fileSize > 10000) {
                    result = result + CellStartRight + (fileSize / 1024) + "&nbsp;KB&nbsp;" + CellEnd;
                } else {
                    result = result + CellStartRight + fileSize + "&nbsp;" + CellEnd;
                }
                //
                result = result + CellStartRightPad + DateString + CellEnd;
                result = result + "</tr>";
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }


        //
        //=================================================================================
        // Returns the Resource Library Row HTML.
        //=================================================================================
        //
        private string GetFormRow_Blank(CPBaseClass cp, string ignore0, string IconLink, string IconOnClick, string Name, string NameLink, DateTime? ModifiedDate, int RowCount, string EditLink, string Description, string FileType, string ResourceName, string ResourceLink, string ImageAlt, string ImageWidth, string ImageHeight, int RecordID, int FolderID, bool AllowEditColumn, bool AllowPlaceColumn, bool AllowSelectColumn, int ColumnCnt) {
            string result = "";
            //
            result = $"\r\n\t<tr class=\"listRow\"><td class=\"left\"><img height=\"23\" width=\"22\" src=\"/ResourceLibrary/spacer.gif\"></td><td class=\"left\" colspan=\"{ColumnCnt - 1}\">&nbsp;</td></tr>";
            //
            return result;
        }
        //
        //=================================================================================
        // Returns the Resource Library Row HTML.
        //=================================================================================
        //
        private string GetFormRow_Options(CPBaseClass cp, int FolderID, string topFolderPath, int ColumnCnt, bool IsContentManagerFiles, bool IsContentManagerFolders, bool hasModifyAccess) {
            string result = "";
            try {
                string FolderCell = "";
                //
                // Inner Cell
                //
                if (hasModifyAccess) {
                    //
                    // if you have viewaccess to the folder, you can see it
                    // if you have modifyaccess to the folder, you can upload to it and create subfolders in it
                    //
                    //If IsContentManagerFolders Then
                    FolderCell = ""
                    + "<table id=\"AddFolderTable\" border=\"0\" cellpadding=\"10\" cellspacing=\"1\" width=\"100%\">"
                    + "<tr>";
                    FolderCell = FolderCell
                    + "<td class=\"left\" align=\"left\" colspan=2>" + kmaAddSpan("Add Folder&nbsp;", "") + "</td>"
                    + "<td class=\"left\" Width=\"99%\" align=\"left\">" + kmaAddSpan("Description&nbsp;", "") + "</td>"
                    + "</tr><tr>"
                    + "<td class=\"left\" Width=\"30\" align=\"right\">1&nbsp;</td>"
                    + "<td class=\"left\" align=\"left\"><INPUT TYPE=\"Text\" NAME=\"FolderName.1\" SIZE=\"30\"></td>"
                    + "<td class=\"left\" align=\"left\"><INPUT TYPE=\"Text\" NAME=\"FolderDescription.1\" SIZE=\"40\"></td>"
                    + "</tr>";
                    FolderCell = FolderCell
                    + "</Table>"
                    + "<table border=\"0\" cellpadding=\"10\" cellspacing=\"1\" width=\"100%\">"
                    + "<tr><td class=\"left\" Width=\"30\"></td><td align=\"left\"><a href=\"#\" onClick=\"InsertFolderRow(); return false;\">+ Add more folders</a></td></tr>"
                    + "</Table>" + htmlHidden("AddFolderCount", 1, "", "AddFolderCount");
                }
                string FileCell = "";
                if (hasModifyAccess) {
                    FileCell = FileCell
                    + "<table id=\"UploadInsert\" border=\"0\" cellpadding=\"0\" cellspacing=\"1\" width=\"100%\">"
                    + "<tr>";
                    FileCell = FileCell
                    + "<td class=\"left\" align=\"left\" colspan=2>" + kmaAddSpan("Add Files&nbsp;", "") + "</td>"
                    + "<td class=\"left\" Width=\"100\" align=\"left\">" + kmaAddSpan("Name&nbsp;", "") + "</td>"
                    + "<td class=\"left\" Width=\"100\" align=\"left\">" + kmaAddSpan("Description&nbsp;", "") + "</td>"
                    + "<td class=\"left\" Width=\"99%\">&nbsp;</td>"
                    + "</tr><tr>"
                    + "<td class=\"left\" Width=\"30\" align=\"right\">1&nbsp;</td>"
                    + "<td class=\"left\" Width=\"200\" align=\"right\"><INPUT TYPE=\"file\" name=\"LibraryUpload.1\"></td>"
                    + "<td class=\"right\" align=\"right\"><INPUT TYPE=\"Text\" NAME=\"LibraryName.1\" SIZE=\"25\"></td>"
                    + "<td class=\"right\" align=\"right\"><INPUT TYPE=\"Text\" NAME=\"LibraryDescription.1\" SIZE=\"39\"></td>"
                    + "<td class=\"left\">&nbsp;</td>"
                    + "</tr>";
                    FileCell = FileCell
                    + "</Table>"
                    + "<table border=\"0\" cellpadding=\"0\" cellspacing=\"1\" width=\"100%\">"
                    + "<tr><td class=\"left\" Width=\"30\"></td><td class=\"left\" align=\"left\"><a href=\"#\" onClick=\"InsertUploadRow(); return false;\">+ Add more files</a></td></tr>"
                    + "</Table>" + htmlHidden("LibraryUploadCount", 1, "", "LibraryUploadCount");
                }
                //
                //
                //
                result = ""
                    + $"<div  style=\"margin-left:10px;\">{cp.Html.CheckBox("AllowThumbnails", cp.User.GetBoolean("LibraryAllowthumbnails", "0"))}&nbsp;Display Thumbnails";
                if (cp.User.IsAdmin || hasModifyAccess) {
                    //
                    string moveSelect = GetMoveFolderPathSelect(cp, FolderID, topFolderPath);
                    if (moveSelect != "") {
                        result += $"<BR>{cp.Html.CheckBox("Move", false)}&nbsp;Move selected files to {moveSelect}";
                    }
                    if (FolderCell != "") {
                        result += $"<hr>{cp.Html.div(FolderCell)}";
                    }
                    if (FileCell != "") {
                        result += $"<hr>{cp.Html.div(FileCell)}";
                    }
                }
                if (result != "") {
                    result = cp.Html.div(result);
                    result = $"<tr><td class=\"bg-light pt-3 left\" colspan={ColumnCnt}>{result}</td></tr>";
                }
                //
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //
        //
        private string GetForm_HeaderCell(CPBaseClass cp, string Align, string Width, string Copy) {
            string Style = ""
                    + "padding: 3px;"
                    + "font-size:10px;";
            string result = $"<td WIDTH=\"{Width}\" ALIGN=\"{Align}\" class=ccAdminListCaption style=\"{Style}\">"
                    + Copy
                    + "</td>";
            return result;
        }
        //
        //
        //
        private bool IsInFolder(CPBaseClass cp, int topFolderID, int FolderID, string ParentPath = "") {
            try {
                if ((FolderID == 0)) {
                    return false;
                }
                if ((topFolderID == 0)) {
                    return true;
                }
                if ((("," + ParentPath + ",").IndexOf("," + FolderID.ToString() + ",", StringComparison.Ordinal) + 1) != 0) {
                    return false;
                }
                ParentPath += "," + FolderID.ToString();
                var folder = DbBaseModel.create<Models.Db.LibraryFolderModel>(cp, FolderID);
                int ParentID = 0;
                if ((folder != null)) {
                    ParentID = folder.parentId;
                }
                if (ParentID == 0) {
                    return false;
                } else if (ParentID == topFolderID) {
                    return true;
                } else {
                    return IsInFolder(cp, topFolderID, ParentID, ParentPath);
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return false;
        }
        //
        //
        //
        private string GetParentFoldersLink(CPBaseClass cp, string topFolderPath, int topFolderID, int currentFolderID, int FolderID, string RefreshQS, string ChildIDList) {
            string result = "";
            try {
                string folderName = "";
                if ((FolderID == 0) || (FolderID == topFolderID)) {
                    //
                    // Root folder
                    folderName = topFolderPath;
                    if (folderName == "") {
                        folderName = "Root";
                    }
                    if (currentFolderID == FolderID) {
                        result = $"Folder <B>{folderName}</B>";
                    } else {
                        result = $"Folder <a href=?{RefreshQS}&FolderID=0>{folderName}</a>";
                    }
                } else {
                    Models.Db.LibraryFolderModel LibraryFolder = DbBaseModel.create<Models.Db.LibraryFolderModel>(cp, $"ID={FolderID}");
                    int ParentID = 0;
                    bool RecordFound = false;
                    if (!(LibraryFolder == null)) {
                        RecordFound = true;
                        ParentID = LibraryFolder.parentId;
                        folderName = LibraryFolder.name;
                    }
                    string FolderLink;
                    //
                    if (currentFolderID == FolderID) {
                        FolderLink = $"<B>{folderName}</B>";
                    } else {
                        FolderLink = $"<a href=?{RefreshQS}&FolderID={FolderID}>{folderName}</a>";
                    }
                    if ((!RecordFound) || (FolderID == topFolderID)) {
                        //
                        // call this the top of the tree
                        if (folderName == "") {
                            folderName = "Root";
                        }
                        result = $"Folder {FolderLink}";
                    } else if ((ChildIDList + ",").IndexOf("," + FolderID + ",", StringComparison.Ordinal) + 1 != 0) {
                        //
                        // circular reference - end it here
                        result = $"Folder (Circular Reference) > {FolderLink}";
                    } else if (currentFolderID == ParentID) {
                        //
                        // circular reference - end it here
                        result = $"Folder {FolderLink}";
                    } else {
                        result = GetParentFoldersLink(cp, topFolderPath, topFolderID, currentFolderID, ParentID, RefreshQS, ChildIDList + "," + FolderID) + "\\" + FolderLink;
                    }
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //----------------------------------------------------------------------------------------
        //   Get a select menu of all folders with which you have ModifyAccess
        //----------------------------------------------------------------------------------------
        //
        private string GetFolderPathSelect(CPBaseClass cp, int topFolderPathID, string topFolderPath, bool RequireModifyAccess) {
            string result = "";
            try {
                string pathRemoveString = "";
                string pathCaption;
                //
                //result = FolderSelect
                if (result == "") {
                    //
                    //
                    // create full paths, set .hasViewAccess
                    //
                    int optionCnt = 0;
                    if (topFolderPath != "") {
                        pathRemoveString = "root\\";
                        int Pos = (topFolderPath.LastIndexOf("\\") + 1);
                        if (Pos > 0) {
                            pathRemoveString = pathRemoveString + topFolderPath.Substring(1 - 1, Pos - 1);
                        }
                    }
                    //
                    // create select
                    //
                    optionCnt = 0;
                    if (topFolderPath == "") {
                        //
                        // if root folder is top folder, everyone has view access
                        //
                        optionCnt = optionCnt + 1;
                        if (topFolderPathID == 0) {
                            //
                            // if root is current folder, mark it selected
                            //
                            result = result + "<option value=0 selected>Root</option>";
                        } else {
                            result = result + "<option value=0>Root</option>";
                        }
                    }
                    int Ptr = FolderPathIndex.GetFirstPointer();
                    while ((Ptr >= 0)) {
                        if (folders[Ptr].hasViewAccess && ((!RequireModifyAccess) || folders[Ptr].hasModifyAccess)) {
                            int PtrFolderID = folders[Ptr].FolderID;
                            pathCaption = folders[Ptr].FullPath.Replace(pathRemoveString, "");

                            if (PtrFolderID == topFolderPathID) {
                                result = result + $"<option value={PtrFolderID} selected>{pathCaption}</option>";
                            } else {
                                result = result + $"<option value={PtrFolderID}>{pathCaption}</option>";
                            }
                            optionCnt = optionCnt + 1;
                        }
                        Ptr = FolderPathIndex.GetNextPointer();
                    }
                    //
                    // Create Select
                    //
                    if (optionCnt <= 1) {
                        //
                        // If only one folder, (the current one), return nothing
                        //
                        result = "";
                    } else {
                        //If result <> "" Then
                        result = $"<select name=FieldName size=1 onChange>{result}</select>";
                    }
                    FolderSelect = result;
                }
                //
                result = FolderSelect;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //
        //
        private string GetFolderPath(CPBaseClass cp, int targetPtr, string ChildIDList) {
            string result = "";
            try {
                //
                int ParentPtr = 0;
                int ParentID = 0;
                int FolderID = 0;

                //
                result = folders[targetPtr].Name;
                ParentID = folders[targetPtr].parentFolderID;
                FolderID = folders[targetPtr].FolderID;
                if (ParentID == 0) {
                    //
                    // At the Root page
                    //
                    result = $"Root\\{result}";
                } else if ((FolderID == ParentID) || (("," + ChildIDList + ",").IndexOf("," + ParentID + ",", StringComparison.Ordinal) + 1 != 0)) {
                    //
                    // circular reference - Make this a root page b
                    //
                } else {
                    for (ParentPtr = 0; ParentPtr <= (folders.Length - 1); ParentPtr++) {
                        //
                        //todo Folder(parentPtr) throws a null ref this needs to be resolved
                        if ((folders[ParentPtr] == null)) {
                            cp.Utils.AppendLogFile($"getfolderPath=6b ******** parentPtr [{ParentPtr}]");
                        } else {
                            if (folders[ParentPtr].FolderID == ParentID) {
                                result = GetFolderPath(cp, ParentPtr, ChildIDList + "," + ParentID) + "\\" + result;
                                //result = GetFolderPath(ParentPtr, ChildIDList & "," & ParentID) & " > " & result
                                break;
                            }
                        }
                    }
                }
                //

            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //
        //
        private string GetJumpFolderPathSelect(CPBaseClass cp, int FolderID, string topFolderPath) {
            string result = "";
            try {
                //
                result = GetFolderPathSelect(cp, FolderID, topFolderPath, false);
                if (result != "") {
                    result = result.Replace("FieldName", "JumpFolderID");
                    result = result.Replace("onChange", "onChange=\"QJump(this);\" ");
                    result = result.Replace("value=", $"value=?{cp.Doc.RefreshQueryString}&FolderID=");
                    result = $"<script language=JavaScript1.2>function QJump(e){{var l=e.value;if(l!=''){{window.name='RL';window.location.assign(l);}}}}</script>{result}";
                }
                //
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;

        }

        //
        //
        //
        private string GetMoveFolderPathSelect(CPBaseClass cp, int FolderID, string topFolderPath) {
            string result = "";
            try {
                //
                result = GetFolderPathSelect(cp, FolderID, topFolderPath, true);
                result = result.Replace("FieldName", "MoveFolderID");
                result = result.Replace("onChange", "onChange=\"var e=getElementById('Move');if(e){e.checked=true};\" ");
                //
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //=============================================================
        //
        //=============================================================
        //
        private string GetRLNav(CPBaseClass cp, int currentFolderID, string topFolderPath, int topFolderID) {
            string result = "";
            bool IsAuthoring;
            try {
                //
                IsAuthoring = false;
                string BakeName = "RLNav";
                if (!IsAuthoring) {
                    //        result = Main.ReadBake(BakeName)
                }
                if (result == "") {
                    string LinkBase = cp.Doc.RefreshQueryString;
                    LinkBase = cp.Utils.ModifyQueryString(LinkBase, "FolderID", "0");

                    //
                    //

                    menuTreeClass Tree = new menuTreeClass(cp);
                    if (topFolderID == 0) {
                        Tree.AddEntry(0.ToString(), (-1).ToString(), "", "", $"?{LinkBase}", "Root");
                    }
                    if (folderCnt > 0) {
                        int Ptr;
                        for (Ptr = 0; Ptr <= folderCnt - 1; Ptr++) {
                            int Id = folders[Ptr].FolderID;
                            if (folders[Ptr].hasViewAccess) {
                                //If hasModifyAccessByFolder(Id, topFolderPath) Then
                                int ParentID = folders[Ptr].parentFolderID;
                                string Caption = folders[Ptr].Name.Replace(" ", "&nbsp;");
                                string Link = $"?{cp.Utils.ModifyQueryString(LinkBase, "FolderID", Id.ToString())}";
                                Tree.AddEntry(Id.ToString(), ParentID.ToString(), "", "", Link, Caption);
                            }
                        }
                    }
                    result = Tree.GetTree(topFolderID.ToString(), currentFolderID.ToString());
                    // Call cp.Response.(BakeName, result, "Library Folders")
                }
                //    //
                //    // Get topFolderPath
                //    //
                //    If topFolderPath = "" Then
                //        topFolderPath = "Root"
                //    Else
                //        topFolderPath = topFolderPath
                //    End If
                //
                // open the current node
                //

                //Call main.AddOnLoadJavascript("convertTrees(); expandToItem('tree0','" & currentFolderID & "');")
                cp.Doc.AddOnLoadJavascript($"convertTrees(); expandToItem('tree0','{currentFolderID}');");
                //Link = "?" & LinkBase
                //Link = "<div style=""position:relative;left:-10;margin-bottom:3px;""><a href=" & Link & " style=""text-decoration:none ! important;"">" & topFolderPath & "</a></div>"
                //result = Replace(result, "<LI ", Link & "<LI ", 1, 1, vbTextCompare)
                ////If CurrentFolderID <> 0 Then
                //result = result & "<script type=""text/javascript"">convertTrees(); expandToItem('tree0','" & CurrentFolderID & "');</script>"
                ////End If
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //
        //
        private bool AllowFolderAccess(CPBaseClass cp, int FolderID, int ParentID) {
            return true;
            ////
            //Try
            //    int GrandParentID
            //    //Dim cs1 As CPCSBaseClass = cp.CSNew()
            //    //
            //    If FolderID = 0 Or cp.User.IsAdmin Then
            //        AllowFolderAccess = True
            //    Else
            //        //AllowFolderAccess = Models.Db.LibraryFolderModel.AllowFolderAccess(cp, FolderID, ParentID)
            //        //
            //        If Not AllowFolderAccess And (ParentID <> 0) Then
            //            Models.Db.LibraryFolderModel LibraryFolder = DbBaseModel.create<Models.Db.LibraryFolderModel>(cp, ParentID)
            //            If Not (LibraryFolder == null) Then
            //                GrandParentID = LibraryFolder.parentId
            //            End If
            //            AllowFolderAccess = AllowFolderAccess(cp, ParentID, GrandParentID)
            //        End If
            //    End If
            //    //
            //Catch ex As Exception
            //    cp.Site.ErrorReport(ex)
            //End Try
        }
        //
        //
        //
        private bool hasModifyAccessByFolder(CPBaseClass cp, int FolderID, string topFolderPath) {
            bool result = false;
            try {
                //
                int Ptr;
                //
                if (FolderID == 86) {
                    FolderID = FolderID;
                }

                //
                if (cp.User.IsAdmin) {
                    //
                    //
                    //
                    result = true;
                } else {
                    //
                    // Need to check permissions
                    //
                    LoadFolders_returnTopFolderId(cp, topFolderPath);
                    if (FolderID == 0) {
                        result = true;
                    } else {
                        Ptr = FolderIdIndex.getPtr(FolderID.ToString());
                        if (Ptr >= 0) {
                            result = folders[Ptr].hasModifyAccess;
                        }
                    }
                }
                //
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //====================================================================================================
        //
        private int LoadFolders_returnTopFolderId(CPBaseClass cp, string topFolderPath) {
            int topFolderID = 0;
            try {
                FolderIdIndex = new FastIndexClass();
                FolderNameIndex = new FastIndexClass();
                FolderPathIndex = new FastIndexClass();
                //
                // Load the folders storage
                List<Models.Db.LibraryFolderModel> foldersList = DbBaseModel.createList<Models.Db.LibraryFolderModel>(cp, "");
                folderCnt = 0;
                if ((foldersList.Count > 0)) {
                    //
                    // Store folders and setup folder index
                    //
                    folderCnt = foldersList.Count;
                    Array.Resize(ref folders, foldersList.Count - 1 + 1);
                    int Ptr = 0;

                    foreach (var folder in foldersList) {
                        folders[Ptr] = new FolderTypeModel();
                        if (true) {
                            if (true) {
                                FolderIdIndex.SetPointer(folder.id.ToString(), Ptr);
                                FolderNameIndex.SetPointer(folder.name, Ptr);
                                folders[Ptr].FolderID = folder.id;
                                folders[Ptr].parentFolderID = folder.parentId;
                                folders[Ptr].Name = folder.name;
                                folders[Ptr].hasModifyAccess = true;
                                folders[Ptr].modifyAccessIsValid = true;
                                folders[Ptr].hasViewAccess = true;
                                //
                                // FullPath, propigate modifyAccess from parent to folder , ViewAccess
                                //
                                //
                                // determine modify access
                                //
                                if ((!folders[Ptr].modifyAccessIsValid)) {
                                    folders[Ptr].hasModifyAccess = LoadFolders_GetModifyAccess(cp, folders[Ptr].parentFolderID);
                                    folders[Ptr].modifyAccessIsValid = true;
                                }
                                //
                                string testFullPath = GetFolderPath(cp, Ptr, "");
                                folders[Ptr].FullPath = testFullPath;
                                //End If
                                FolderPathIndex.SetPointer(testFullPath, Ptr);
                                //
                                if (($"root\\{topFolderPath}").IndexOf(testFullPath, StringComparison.OrdinalIgnoreCase) == 0) {
                                    //
                                    folders[Ptr].hasViewAccess = true;
                                }
                                //
                                //
                                topFolderID = 0;
                                if (topFolderPath != "") {
                                    string[] targetFolders = topFolderPath.Split(new string[] { "\\" }, StringSplitOptions.None);
                                    int targetFolderCnt = (targetFolders.Length - 1) + 1;
                                    topFolderID = loadFolders_getFolderID(cp, targetFolders, targetFolderCnt - 1);
                                    //
                                    // if topFolderId not found, create the new folder(s) necessary to targetFolderPath
                                    //
                                    string targetFolderName = "";
                                    if (topFolderID == 0) {
                                        int targetFolderId = 0;
                                        for (Ptr = 0; Ptr <= targetFolderCnt - 1; Ptr++) {
                                            targetFolderName = targetFolders[Ptr];
                                            int targetParentFolderID = targetFolderId;
                                            //
                                            // find or create the folder with this name and this targetParentFolderID
                                            //
                                            int testFolderPtr = FolderNameIndex.getPtr(targetFolders[Ptr]);
                                            while (testFolderPtr >= 0) {
                                                int testParentID = folders[testFolderPtr].parentFolderID;
                                                if (targetParentFolderID != testParentID) {
                                                    //
                                                    // right name but wrong parent, try next
                                                    //
                                                } else {
                                                    //
                                                    // good match, this as the parent and find the next
                                                    //
                                                    break;
                                                }
                                                testFolderPtr = FolderNameIndex.GetNextPointerMatch(targetFolderName);
                                            }
                                            if (testFolderPtr >= 0) {
                                                targetFolderId = folders[testFolderPtr].FolderID;
                                            } else {
                                                //
                                                // folder not found, create it with the parent
                                                //
                                                //cS = main.InsertCSRecord("Library Folders")

                                                //If main.IsCSOK(cS) Then
                                                if (!(folder == null)) {
                                                    targetFolderId = folder.id;
                                                    folder.name = targetFolderName;
                                                    folder.parentId = targetParentFolderID;
                                                    //Call main.SetCS(cS, "name", targetFolderName)
                                                    //Call main.SetCS(cS, "parentid", targetParentFolderID)
                                                    folder.save(cp);
                                                }
                                                //Call main.CloseCS(cS)
                                            }
                                            if (Ptr == (targetFolderCnt - 1)) {
                                                topFolderID = targetFolderId;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        Ptr += 1;
                    }
                    //
                }
                //
                topFolderID = topFolderID;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return topFolderID;
        }

        //Private Function IsEmpty(folderCells As Object) As Boolean
        //    Throw New NotImplementedException()
        //End Function
        //
        //===============================================================================================
        //   returns the id of the cache folder that matches the target folder
        //       targetfolder = 'tier1\tier2\tier3'
        //       targetArray=['tier1','tier2','tier3'], targetArray(0)='tier1'
        //       targetArrayPtr is the index into the targetArray of the folder we are looking up
        //       returns the id of the folder 'tier3' that has a parent folder 'tier2', etc.
        //       if not folder exists, it returns 0
        //===============================================================================================
        //
        private int loadFolders_getFolderID(CPBaseClass cp, string[] targetArray, int targetArrayPtr) {
            int result = 0;
            //
            int cachePtr;
            int cacheFolderID;
            int cacheParentFolderID;
            string targetFolderName;
            int targetFolderParentId;
            //
            result = 0;
            targetFolderName = targetArray[targetArrayPtr];
            cachePtr = FolderNameIndex.getPtr(targetFolderName);
            while (cachePtr >= 0) {
                cacheFolderID = folders[cachePtr].FolderID;
                if (targetArrayPtr == 0) {
                    //
                    // this was the top-most folder, return the non-zero cache id
                    //
                    if (folders[cachePtr].parentFolderID != 0) {
                        //
                        // top of target path but record parent <> 0, try next record
                        //
                    } else {
                        //
                        // top of target path matches records (parentid=0)
                        //
                        result = cacheFolderID;
                        break;
                    }
                } else {
                    //
                    // not top-most, since there could be multiple matches, test the parent
                    // of this match, if it is ok, return with this id. If not, try the next
                    // cache match for this folder name.
                    //
                    targetFolderParentId = loadFolders_getFolderID(cp, targetArray, targetArrayPtr - 1);
                    if (targetFolderParentId <= 0) {
                        //
                        // parent folder not found, try the next matching folder name
                        //
                    } else {
                        //
                        // parent folder found, check that its target folder matches the parent id of this folder
                        //
                        if (targetFolderParentId == folders[cachePtr].parentFolderID) {
                            //
                            // this folder is correct, return with it's ID
                            //
                            result = cacheFolderID;
                            break;
                        } else {
                            //
                            // the cache folder hierarchy does not match the traget folder string, try next name match
                            //
                        }
                    }
                }
                cachePtr = FolderNameIndex.GetNextPointerMatch(targetFolderName);
            }
            //
            return result;
        }
        //
        //
        //
        private bool LoadFolders_GetModifyAccess(CPBaseClass cp, int FolderID) {
            bool result = false;
            //
            int Ptr;
            //
            Ptr = FolderIdIndex.getPtr(FolderID.ToString());
            if (Ptr >= 0) {
                if (folders[Ptr].modifyAccessIsValid) {
                    //
                    //
                    //
                    result = folders[Ptr].hasModifyAccess;
                } else if (folders[Ptr].parentFolderID == 0) {
                    //
                    // Parent is root, this folder does not have access
                    //
                    result = false;
                    folders[Ptr].hasModifyAccess = result;
                    folders[Ptr].modifyAccessIsValid = true;
                } else {
                    //
                    // Parent is not root
                    //
                    result = LoadFolders_GetModifyAccess(cp, folders[Ptr].parentFolderID);
                    folders[Ptr].hasModifyAccess = result;
                    folders[Ptr].modifyAccessIsValid = true;
                }
            }
            //
            return result;
        }
        //
        //=================================================================================
        // Create
        //=================================================================================
        //
        public LegacyLibraryClass() {
            iMinRows = 10;
        }
        //
        //=================================================================================
        // Kill - removed, no longer needed in C#
        //=================================================================================
        //
        //
        //
        private int GetFileSize(CPBaseClass cp, string VirtualFilePathPage) {
            string hint;
            int result = 0;
            try {
                //
                hint = "1";
                //
                hint = "2";
                VirtualFilePathPage = VirtualFilePathPage.Replace("/", "\\");
                int SlashPosition = (VirtualFilePathPage.LastIndexOf("\\") + 1);
                string Filename = "";
                string Pathname = "";
                if (SlashPosition != 0) {
                    Filename = VirtualFilePathPage.Substring(SlashPosition - 1 + 1).ToLower();
                    Pathname = VirtualFilePathPage.Substring(1 - 1, SlashPosition - 1);
                }
                //
                string FileDescriptor;
                FileDescriptor = cp.File.fileList(Pathname);
                hint = "3";
                if (FileDescriptor == "") {
                    //Call AppendLogFile("GetFileSize, descriptor is blank")
                } else {
                    hint = "4";
                    string[] FileSplit2 = FileDescriptor.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                    //Call AppendLogFile("GetFileSize, FileDescriptor lines=" & UBound(FileSplit2))
                    hint = "5";
                    int Ptr;
                    for (Ptr = 0; Ptr <= (FileSplit2.Length - 1); Ptr++) {
                        string[] FileParts = FileSplit2[Ptr].Split(new string[] { "\t" }, StringSplitOptions.None);
                        if ((FileParts.Length - 1) <= 5) {
                            //Call AppendLogFile("GetFileSize, FileDescriptor row [" & Ptr * "] has <6 parts, descrriptor=" & FileDescriptor)
                        } else {
                            if (FileParts[0].ToLower() == Filename) {
                                result = cp.Utils.EncodeInteger(FileParts[5]);
                                //Call AppendLogFile("GetFileSize, match on " & FileParts(0))
                                break;
                            }
                        }
                    }
                    hint = "6";
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        //
        //
        //
        private int GetFileTypeID(CPBaseClass cp, string Filename) {
            int result = 0;
            //
            string[] FileNameSplit;
            string FileExtension;
            int CSType;
            int DefaultFileTypeID = 0;
            int cnt;
            int Ptr;
            string hint;
            //
            FileNameSplit = Filename.Split(new string[] { "." }, StringSplitOptions.None);
            FileExtension = FileNameSplit[(FileNameSplit.Length - 1)];
            //
            // try to read if from IconFiles
            //
            hint = "1";
            result = DefaultFileTypeID;
            foreach (FileTypeModel iconFile in iconFiles) {
                hint = "2";
                if ((("," + iconFile.ExtensionList + ",").IndexOf("," + FileExtension + ",", StringComparison.OrdinalIgnoreCase) + 1) != 0) {
                    hint = "3";
                    result = iconFile.FileTypeID;
                    break;
                }
                if (iconFile.Name.ToLower() == "default") {
                    hint = "4";
                    DefaultFileTypeID = iconFile.FileTypeID;
                }
            }
            //cnt = IconFileCnt
            //If cnt > 0 Then
            //    For Ptr = 0 To cnt - 1
            //    Next
            //End If
            hint = "5";
            //    //
            //    // try Db next
            //    //
            //    If result = 0 Then
            //hint = "6"
            //        CSType = Main.OpenCSContent("Library File Types", "(extensionlist like '%," & FileExtension & ",%')or(extensionlist like '%,." & FileExtension & ",%')")
            //        If Main.IsCSOK(CSType) Then
            //            result = Main.GetCSInteger(CSType, "ID")
            //        End If
            //        Call Main.closecs(CSType)
            //        If result = 0 Then
            //            result = Main.GetRecordID("Library File Types", "default")
            //        End If
            //    End If
            //
            return result;
        }


    }
}
