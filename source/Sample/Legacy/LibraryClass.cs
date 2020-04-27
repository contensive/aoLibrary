using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using System.Net;
using System.Web;
using Contensive.Addons.ResourceLibrary.Controllers;
using Contensive.BaseClasses;
using Contensive.Addons.ResourceLibrary.Controllers.genericController;
using Contensive.Addons.ResourceLibrary.Models;
using Contensive.VbConversion;

namespace Contensive.Addons.ResourceLibrary.Views {
    // 
    public class LibraryClass : AddonBaseClass {
        // 
        // Private main As Contensive.vbConversion.MainClass
        public override object Execute(CPBaseClass CP) {
            string returnHtml = "";
            try {
                // 
                returnHtml = GetContent(CP);
            }
            // 
            catch (Exception ex) {
                CP.Site.ErrorReport(ex);
            }
            return returnHtml;
        }
        // 
        // 
        public List<FileType> iconFiles = new List<FileType>();
        // Public IconFiles() As FileType
        // Public IconFileCnt As Integer
        // 
        // 
        // 
        public class FolderType {
            public int FolderID;
            public int parentFolderID;
            public string Name;
            public string FullPath;
            // 
            public bool hasViewAccess;                    // has permission to view this folder (below topFolderPath)
            public bool viewAccessIsValid;                 // true when hasViewAccess is correct
            // 
            public bool hasModifyAccess;                  // has permission to modify files and folders in this folder
            public bool modifyAccessIsValid;              // true when hasModifyAccess is correct
        }
        public FolderType[] folders = new[] { };
        public int folderCnt;
        public VbConversion.FastIndexClass FolderIdIndex = new VbConversion.FastIndexClass();
        public VbConversion.FastIndexClass FolderNameIndex = new VbConversion.FastIndexClass();
        public VbConversion.FastIndexClass FolderPathIndex = new VbConversion.FastIndexClass();
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
        // This means the resource library supports buttons that allow objects to be
        // placed on different page from the resource library, like an Editor
        // 
        public bool AllowPlace;
        // 
        // ----- If an editor is used to call the resource library, the window.opener.insertresource()
        // call needs the object name of the editor so the contents can be copied to the invisible
        // form field after the changes (no onchange event available)
        // 
        public string SelectResourceEditorObjectName;
        // 
        // ----- If AllowPlace is true and SelectLinkObjectName<>"", the RL is being used as a link selector
        // When the 'place' icon is clicked, the URL of the resource is copied to the window.opener.[selectlinkobjectname]
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
        // ' SourceMode
        // '   3/6/2010 - moved codes up to capture the 0 case and it to page
        // '   1 = From Editor Object or Link selector: allow image and download insert, provide close button
        // '   2 = From Editor Image Properties: allow image insert, provide close button
        // '   3 = From Admin site, no inserts, and provide cancel button
        // Const SourceModeOnPage = 1
        // Const SourceModeFromDownloadRequest = 2
        // Const SourceModeFromLinkDialog = 3
        // 0 = From Editor Object selector: allow image and download insert, provide close button
        // 1 = From Editor Image Properties: allow image insert, provide close button
        // 2 = From Admin site, no inserts, and provide cancel button
        public const string SourceModeFromDownloadRequest = 0;
        public const string SourceModeFromLinkDialog = 1;
        public const string SourceModeOnPage = 2;
        // 
        // 0 caller is the editor directly, clicking on action icons calls InsertImaage, etc
        // 1 caller is the editor image page, clicking on action icons calls the image page methods
        // 
        public int HoldPosition;
        // 
        //====================================================================================================
        // 
        public string GetContent(CPBaseClass cp) {
            string result = "";
            try {
                string topFolderPath = "";
                bool AllowGroupAdd;
                string OptionString = "";
                // 
                topFolderPath = cp.Doc.GetText("RootFolderName");
                cp.Site.TestPoint("topFolderPath=[" + topFolderPath + "]");
                AllowGroupAdd = cp.Utils.EncodeBoolean(cp.Doc.GetBoolean("AllowGroupAdd"));
                AllowPlace = cp.Utils.EncodeBoolean(cp.Doc.GetBoolean("AllowSelectResource"));
                SelectResourceEditorObjectName = cp.Doc.GetText("SelectResourceEditorObjectName");
                SelectLinkObjectName = cp.Doc.GetText("SelectLinkObjectName");
                blockFolderNavigation = cp.Utils.EncodeBoolean(cp.Doc.GetBoolean("Block Folder Navigation"));
                // 
                // topFolder should be in this format toptier\tier2\tier2
                // all lowercase, no leading or trailing slashes, backslashs, remove 'root\'
                // 
                topFolderPath = Strings.Trim(topFolderPath);
                topFolderPath = Strings.LCase(topFolderPath);
                topFolderPath = Strings.Replace(topFolderPath, "/", @"\");
                if (Strings.Left(topFolderPath, 4) == "root")
                    topFolderPath = Strings.Mid(topFolderPath, 5);
                if (Strings.Left(topFolderPath, 1) == @"\")
                    topFolderPath = Strings.Mid(topFolderPath, 2);
                if (Strings.Right(topFolderPath, 1) == @"\")
                    topFolderPath = Strings.Mid(topFolderPath, 1, Strings.Len(topFolderPath) - 1);
                // 
                GetContent = GetForm(cp, topFolderPath, AllowGroupAdd);
                result = GetContent;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        // 
        //====================================================================================================
        // 
        private string GetForm(CPBaseClass cp, string topFolderPath, bool AllowGroupAdd) {
            string result = "";
            try {
                const string LibraryFileTypespathFilename = @"resourcelibrary\LibraryConfig.xml";
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
                    // BuildVersion = cp.Site.GetText("build version")
                    bool IsContentManagerFiles = cp.User.IsContentManager("Library Files");
                    bool IsContentManagerFolders = cp.User.IsContentManager("Library Folders");
                    string Button = cp.Doc.GetText("Button");
                    string FileTypeFilter = LCase(cp.Doc.GetText("ffilter"));
                    cp.Doc.AddRefreshQueryString("ffilter", FileTypeFilter);
                    bool AllowThumbnails = cp.User.GetBoolean("LibraryAllowthumbnails", "0");
                    string FolderIDString = cp.Doc.GetText("folderid");
                    int currentFolderID = cp.Utils.EncodeInteger(FolderIDString);
                    if (FolderIDString != "")
                        cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString());
                    else
                        currentFolderID = cp.User.GetInteger("Libraryfolderid", "0");
                    // 
                    // Load Folder cache
                    // 
                    hint = "010, topFolderPath=" + topFolderPath;
                    int topFolderID = LoadFolders_returnTopFolderId(cp, topFolderPath);
                    // 
                    bool reloadFolderCache = false;
                    int currentFolderPtr;
                    // 
                    // verify that current folder has viewAccess (if not jumpt to root)
                    // 
                    if (currentFolderID != 0) {
                        currentFolderPtr = FolderIdIndex.getPtr(currentFolderID.ToString());
                        if ((currentFolderPtr > Information.UBound(folders)) | (currentFolderPtr < 0))
                            currentFolderPtr = 0;
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
                    if ((cp.User.IsAdmin | IsContentManagerFiles | IsContentManagerFolders))
                        // 
                        // you get modify access if you can modify the content
                        // 
                        currentFolderHasModifyAccess = true;
                    else if (currentFolderID == 0) {
                    } else {
                        // 
                        // others have modify access to this folder if they are in a modify access group
                        // 
                        currentFolderPtr = FolderIdIndex.getPtr(System.Convert.ToString(currentFolderID));
                        if (currentFolderPtr >= 0)
                            currentFolderHasModifyAccess = folders[currentFolderPtr].hasModifyAccess;
                    }
                    // topFolderID = GetFolderID(topFolderPath)
                    // 
                    // Load IconFiles
                    // 
                    hint = "030";
                    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                    doc.LoadXml(cp.File.ReadVirtual(LibraryFileTypespathFilename));
                    int Ptr;
                    hint = "040";
                    if ((doc.DocumentElement.Name.ToLower().Equals("libraryconfig"))) {
                        if (doc.DocumentElement.ChildNodes.Count > 0) {
                            {
                                var withBlock = doc.DocumentElement;
                                Ptr = 0;
                                hint = "050";
                                System.Xml.XmlElement baseNode;
                                foreach (var baseNode in withBlock.ChildNodes) {
                                    hint = "060";
                                    switch (Strings.LCase(baseNode.Name)) {
                                        case "filetype": {
                                                hint = "070";
                                                Ptr = Ptr + 1;
                                                FileType newFileType = new FileType();
                                                iconFiles.Add(newFileType);
                                                // Dim IconCnt As Integer
                                                // If Ptr >= IconCnt Then
                                                // IconCnt = IconCnt + 10
                                                // ReDim Preserve IconFiles(IconCnt)
                                                // End If
                                                {
                                                    var withBlock1 = newFileType;
                                                    System.Xml.XmlElement typeNode;
                                                    foreach (var typeNode in baseNode.ChildNodes) {
                                                        switch (Strings.LCase(typeNode.Name)) {
                                                            case "name": {
                                                                    withBlock1.Name = typeNode.Value;
                                                                    break;
                                                                }

                                                            case "filetypeid": {
                                                                    withBlock1.FileTypeID = cp.Utils.EncodeInteger(typeNode.Value);
                                                                    break;
                                                                }

                                                            case "extensionlist": {
                                                                    withBlock1.ExtensionList = typeNode.Value;
                                                                    break;
                                                                }

                                                            case "isdownload": {
                                                                    withBlock1.IsDownload = cp.Utils.EncodeBoolean(typeNode.Value);
                                                                    break;
                                                                }

                                                            case "isimage": {
                                                                    withBlock1.IsImage = cp.Utils.EncodeBoolean(typeNode.Value);
                                                                    break;
                                                                }

                                                            case "isvideo": {
                                                                    withBlock1.IsVideo = cp.Utils.EncodeBoolean(typeNode.Value);
                                                                    break;
                                                                }

                                                            case "isflash": {
                                                                    withBlock1.IsFlash = cp.Utils.EncodeBoolean(typeNode.Value);
                                                                    break;
                                                                }

                                                            case "iconlink": {
                                                                    withBlock1.IconFilename = typeNode.Value;
                                                                    break;
                                                                }

                                                            case "mediaiconlink": {
                                                                    withBlock1.MediaIconFilename = typeNode.Value;
                                                                    break;
                                                                }

                                                            case "downloadiconlink": {
                                                                    withBlock1.DownloadIconFilename = typeNode.Value;
                                                                    break;
                                                                }
                                                        }
                                                    }
                                                }

                                                break;
                                            }
                                    }
                                }
                            }
                        }
                    }
                    // 
                    // Verify default icons
                    // 
                    hint = "100";
                    string DefaultIcon = @"\cclib\images\IconImage2.gif";
                    string DefaultMedia = @"\cclib\images\Iconimage2Media.gif";
                    string DefaultDownload = @"\cclib\images\Iconimage2Download.gif";
                    // 
                    if (cp.Doc.GetText("SourceMode") == "")
                        SourceMode = SourceModeOnPage;
                    else
                        SourceMode = cp.Doc.GetInteger("SourceMode");
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
                    LibraryFolderModel folder = LibraryFolderModel.create(cp, currentFolderID);
                    // Dim FolderGroupName As String
                    int FolderParentID;
                    if ((folder != null))
                        FolderParentID = folder.ParentID;
                    if ((topFolderID != currentFolderID) & (topFolderID != FolderParentID)) {
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
                    bool AllowPlaceColumn = AllowPlace & ((SourceMode == SourceModeFromLinkDialog) | (SourceMode == SourceModeFromDownloadRequest));
                    if (AllowPlaceColumn)
                        ColumnCnt = ColumnCnt + 1;
                    bool AllowEditColumn = (IsContentManagerFiles | IsContentManagerFolders);
                    if (AllowEditColumn)
                        ColumnCnt = ColumnCnt + 1;
                    bool AllowSelectColumn = currentFolderHasModifyAccess;
                    if (AllowSelectColumn)
                        ColumnCnt = ColumnCnt + 1;
                    // 
                    // ----- Setup folder editing
                    bool AllowFolderAuthoring = IsContentManagerFolders;
                    int FolderCID;
                    if (AllowFolderAuthoring)
                        FolderCID = cp.Content.GetID("Library Folders");
                    // 
                    // ----- Setup file editing
                    bool AllowFileAuthoring = IsContentManagerFiles;
                    int FileCID;
                    if (AllowFileAuthoring)
                        FileCID = cp.Content.GetID("Library Files");
                    // Dim FolderGroupID as integer
                    // 
                    // ----- Setup Local File Management
                    // Allow if Content Manager or user has group membership
                    // Always allow, everyone has access to the root folder, then if you can get to the folder, let em upload
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
                            case object _ when ButtonCancel: {
                                    // 
                                    // CAncel button, just redirect back to the current page
                                    cp.Response.Redirect("#");
                                    break;
                                }

                            case object _ when ButtonDelete: {
                                    // 
                                    // 
                                    // 
                                    RowCount = cp.Doc.GetInteger("RowCount");
                                    int DeleteFileID;
                                    int DeleteFolderID;
                                    if (RowCount > 0) {
                                        for (Ptr = 0; Ptr <= RowCount - 1; Ptr++) {
                                            if (cp.Doc.GetBoolean("Row" + Ptr)) {
                                                DeleteFolderID = cp.Doc.GetInteger("Row" + Ptr + "FolderID");
                                                if (DeleteFolderID != 0) {
                                                    // Call Main.WriteStream("Deleting Folder " & FolderID)
                                                    cp.Content.Delete("Library Folders", "id=" + DeleteFolderID);
                                                    reloadFolderCache = true;
                                                }
                                                DeleteFileID = cp.Doc.GetInteger("Row" + Ptr + "FileID");
                                                if (DeleteFileID != 0) {
                                                    // Call Main.WriteStream("Deleting File " & FileID)
                                                    cp.Content.Delete("Library Files", "id=" + DeleteFileID);
                                                    reloadFolderCache = true;
                                                }
                                            }
                                        }
                                    }

                                    break;
                                }

                            case object _ when ButtonApply: {
                                    // 
                                    // Move Files
                                    // 
                                    if (cp.Doc.GetBoolean("Move")) {
                                        int targetFolderId = cp.Doc.GetInteger("MoveFolderID");
                                        RowCount = cp.Doc.GetInteger("RowCount");
                                        if (RowCount > 0) {
                                            for (Ptr = 0; Ptr <= RowCount - 1; Ptr++) {
                                                if (cp.Doc.GetBoolean("Row" + Ptr)) {
                                                    int MoveFolderID = cp.Doc.GetInteger("Row" + Ptr + "FolderID");
                                                    int MoveFileID = cp.Doc.GetInteger("Row" + Ptr + "FileID");
                                                    if (MoveFolderID != 0) {
                                                        cp.Db.ExecuteSQL("update ccLibraryFolders set ParentID=" + targetFolderId + " where ID=" + MoveFolderID);
                                                        reloadFolderCache = true;
                                                    } else if (MoveFileID != 0) {
                                                        cp.Db.ExecuteSQL("update ccLibraryFiles set FolderID=" + targetFolderId + " where ID=" + MoveFileID);
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
                                            string folderName = cp.Doc.GetText("FolderName." + UploadPointer);
                                            if (folderName != "") {
                                                if (IsContentManagerFolders & (!cp.User.IsAdmin) & (currentFolderID == 0)) {
                                                    // 
                                                    // Content Managers can not add folders to the root folder
                                                    // 
                                                    cp.UserError.Add("Your account does not have access to add new folders to the root folder.");
                                                    break;
                                                } else {
                                                    Models.LibraryFolderModel libraryFolder = Models.LibraryFolderModel.add(cp);
                                                    libraryFolder.name = folderName;
                                                    libraryFolder.Description = cp.Doc.GetText("FolderDescription." + UploadPointer);
                                                    libraryFolder.ParentID = currentFolderID;
                                                    libraryFolder.save(cp);
                                                    // cS = Main.InsertCSRecord("Library Folders")
                                                    // If Main.IsCSOK(cS) Then
                                                    // Copy = cp.Doc.GetText("FolderDescription." & UploadPointer)
                                                    // Call Main.SetCS(cS, "Name", folderName)
                                                    // Call Main.SetCS(cS, "Description", Copy)
                                                    // If currentFolderID <> 0 Then
                                                    // Call Main.SetCS(cS, "ParentID", currentFolderID)
                                                    // End If
                                                    // End If
                                                    // Call Main.closecs(cS)
                                                    reloadFolderCache = true;
                                                }
                                            }
                                        }
                                        // 
                                        // Upload files
                                        // 
                                        hint = "400";
                                        int UploadCount = cp.Doc.GetInteger("LibraryUploadCount");
                                        string ImageFilename = "";
                                        // Dim imagefileFolderId As Integer = cp.Doc.GetInteger("FolderID")
                                        for (UploadPointer = 1; UploadPointer <= UploadCount; UploadPointer++) {
                                            string imageRequestName = RequestNameLibraryUpload + "." + UploadPointer;
                                            ImageFilename = cp.Doc.GetText(RequestNameLibraryUpload + "." + UploadPointer);
                                            if (ImageFilename != "") {
                                                hint = "410";
                                                LibraryFileModel libraryFile = LibraryFileModel.add(cp);


                                                string libraryName = cp.Doc.GetText(RequestNameLibraryName + "." + UploadPointer);
                                                if (libraryName == "")
                                                    libraryName = ImageFilename;
                                                libraryFile.name = libraryName;
                                                var libraryDescription = cp.Doc.GetText(RequestNameLibraryDescription + "." + UploadPointer);
                                                if (libraryDescription == "")
                                                    libraryDescription = ImageFilename;
                                                FileExtension = "";
                                                FilenameNoExtension = "";
                                                AltSizeList = "";
                                                Pos = Strings.InStrRev(ImageFilename, ".");
                                                if (Pos > 0) {
                                                    FileExtension = Strings.Mid(ImageFilename, Pos + 1);
                                                    FilenameNoExtension = Strings.Left(ImageFilename, Pos - 1);
                                                }
                                                // ''''libraryFile.Filename.upload(cp, imageRequestName)

                                                string VirtualFilePathPage = libraryFile.getUploadPath("filename");


                                                string VirtualFilePath = Strings.Replace(VirtualFilePathPage, ImageFilename, "");
                                                libraryFile.Description = libraryDescription;
                                                libraryFile.FolderID = currentFolderID;
                                                cp.Html.ProcessInputFile(imageRequestName, VirtualFilePath);

                                                libraryFile.FileSize = GetFileSize(cp, cp.Site.PhysicalFilePath + libraryFile.name);
                                                string FileType = "";
                                                hint = "425";
                                                FileTypeID = getFileTypeID(cp, ImageFilename);
                                                libraryFile.FileTypeID = FileTypeID;
                                                libraryFile.name = libraryName;
                                                libraryFile.Description = libraryDescription;
                                                libraryFile.Filename = VirtualFilePath + ImageFilename;
                                                libraryFile.ModifiedDate = DateTime.Now;
                                                libraryFile.save(cp);

                                                reloadFolderCache = true;
                                            }
                                        }
                                    }

                                    break;
                                }
                        }
                    }
                    hint = "500";
                    if (reloadFolderCache) {
                        folderCnt = 0;
                        topFolderID = LoadFolders_returnTopFolderId(cp, topFolderPath);
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
                    if ((SourceMode == SourceModeFromDownloadRequest) | (SourceMode == SourceModeFromLinkDialog))
                        ButtonExit = cp.Html.Button(rnbutton, ButtonClose, "", "windowcloseID");
                    else
                        ButtonExit = cp.Html.Button(rnbutton, ButtonCancel);
                    string ButtonBar = "";
                    if (AllowLocalFileAdd) {
                        if (currentFolderHasModifyAccess)
                            ButtonBar = "<div class=ccAdminButtonBar>"
+ ButtonExit
+ cp.Html.Button(rnbutton, ButtonApply)
+ cp.Html.Button(rnbutton, ButtonDelete, RequestNameButton, "returnDeleteCheckID")
+ "</div>";
                        else
                            ButtonBar = "<div class=ccAdminButtonBar>"
                                + cp.Html.Button(rnbutton, ButtonApply)
                                + "</div>";
                    }

                    // result = result & genericController.htmlHidden("FolderID", currentFolderID) 
                    result = result + ButtonBar;

                    string JumpSelect = "";
                    JumpSelect = GetJumpFolderPathSelect(cp, currentFolderID, topFolderPath);
                    result = result + "<div style=\"padding:10px;\">" + GetParentFoldersLink(cp, topFolderPath, topFolderID, currentFolderID, currentFolderID, cp.Doc.RefreshQueryString, "") + "</div>";
                    if (JumpSelect != "")
                        result = result + "<div style=\"padding:10px;padding-top:0px\">" + "Jump to&nbsp;" + JumpSelect + "</div>";
                    // 
                    // From here down the form divides into FormFolder and FormDetails
                    // 
                    string FormDetails = "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\"><tr class=\"headRow\">";
                    if (AllowSelectColumn)
                        FormDetails = FormDetails + GetForm_HeaderCell(cp, "center", "10", "Select<BR>" + spacer1x10);
                    if (AllowEditColumn)
                        FormDetails = FormDetails + GetForm_HeaderCell(cp, "center", "15", "Edit<br>" + spacer1x15);
                    if (AllowPlaceColumn)
                        FormDetails = FormDetails + GetForm_HeaderCell(cp, "center", "15", "Place<br>" + spacer1x15);
                    FormDetails = FormDetails
                            + GetForm_HeaderCell(cp, "left", "20", "&nbsp;<BR>" + spacer1x20)
                            + GetForm_HeaderCell(cp, "left", "20%", "Name<br>" + spacer1x20)
                            + GetForm_HeaderCell(cp, "left", "50%", "Description<br>" + spacer1x15)
                            + GetForm_HeaderCell(cp, "center", "50", "Size<br>" + spacer1x50)
                            + GetForm_HeaderCell(cp, "center", "50", "Modified&nbsp;&nbsp;<br>" + spacer1x50)
                            + "</tr>";
                    // 
                    // ----- Select the Folder Rows
                    // 
                    string Criteria = "((ParentID is null)or(ParentID=0))";
                    // 
                    if (currentFolderID != 0)
                        cp.Doc.AddRefreshQueryString("FolderID", currentFolderID.ToString());
                    // 
                    string SortField = cp.Doc.GetText("sortfield");
                    if (SortField == "")
                        SortField = "Name";
                    cp.Doc.AddRefreshQueryString("SortField", SortField);
                    // 
                    int SortDirection = cp.Doc.GetInteger("sortdirection");
                    if (SortDirection != 0)
                        cp.Doc.AddRefreshQueryString("SortDirection", SortDirection.ToString());
                    // 
                    if (SortDirection != 0 & SortField != "")
                        SortField = SortField + " DESC";
                    // 
                    LibraryFolderModel parentFolder = null/* TODO Change to default(_) if this is not a reference type */;

                    int parentFolderID;
                    if (currentFolderID != 0) {
                        // 
                        // ----- FolderID given, lookup record and get ParentID
                        // Note that allowupfolder allows users to "up" past top if they set it manually
                        // Fix this when security is added
                        // 
                        folder = LibraryFolderModel.create(cp, currentFolderID);
                        if ((folder != null))
                            parentFolderID = folder.ParentID;
                        parentFolder = LibraryFolderModel.create(cp, parentFolderID);
                        Criteria = "(ParentID=" + KmaEncodeSQLNumber(cp, currentFolderID) + ")";
                    } else if (topFolderPath != "") {
                        // 
                        // ----- Rootfolder given, lookup record and get ParentID
                        // 
                        folder = LibraryFolderModel.createByName(cp, topFolderPath);
                        if ((folder != null)) {
                            parentFolderID = 0;
                            currentFolderID = folder.id;
                            cp.User.SetProperty("LibraryFolderID", currentFolderID.ToString());
                        }
                        parentFolder = LibraryFolderModel.create(cp, parentFolderID);
                        Criteria = "(ParentID=" + KmaEncodeSQLNumber(cp, currentFolderID) + ")";
                    } else
                        // 
                        // ----- Use Root as top (no record)
                        parentFolder = LibraryFolderModel.create(cp, parentFolderID);
                    // 
                    // ----- Output the page
                    // 
                    RowCount = 0;
                    hint = "700";
                    if (true) {
                        // 
                        // ----- List out the folders
                        List<LibraryFolderModel> folderList = LibraryFolderModel.createList(cp, Criteria, SortField);
                        string IconLink;
                        string EditLink;
                        DateTime ModifiedDate;
                        foreach (var folder in folderList) {
                            string ChildFolderName = folder.name;
                            if (ChildFolderName == "")
                                ChildFolderName = "[no name]";
                            EditLink = "";
                            if (AllowFolderAuthoring)
                                EditLink = adminUrl(cp) + "?cid=" + FolderCID + "&id=" + folder.id + "&af=4" + "&aa=2&depth=1";
                            IconLink = cp.Utils.ModifyQueryString(cp.Doc.RefreshQueryString, "folderid", System.Convert.ToString(folder.id));
                            ModifiedDate = folder.ModifiedDate;
                            if (ModifiedDate <= DateTime.MinValue)
                                ModifiedDate = folder.DateAdded;
                            int ChildFolderID;
                            FormDetails = FormDetails + GetFormRow_ChildFolders(cp, IconFolderClosed, IconLink, "", ChildFolderName, "", ModifiedDate, RowCount, EditLink, folder.Description, "CHILD", "", "", "", "", "", 0, folder.id, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn);
                            RowCount = RowCount + 1;
                        }

                        // 
                        // Lookup the files in the folder
                        // 
                        hint = "720";
                        if (currentFolderID == 0)
                            Criteria = "((FolderID is null)or(FolderID=0))";
                        else
                            Criteria = "(FolderID=" + KmaEncodeSQLNumber(cp, currentFolderID) + ")";
                        // FieldList = "ID,Name,ModifiedDate,Filename,Width,Height,DateAdded,Description,AltText,FileTypeID,FileSize,AltSizeList"
                        if (currentFolderID == 0)
                            Criteria = "((FolderID is null)or(FolderID=0))";
                        else
                            Criteria = "(FolderID=" + KmaEncodeSQLNumber(cp, currentFolderID) + ")";
                        List<LibraryFileModel> fileList = LibraryFileModel.createList(cp, Criteria, SortField);
                        foreach (var file in fileList) {
                            bool UpdateRecord = false;
                            int ResourceRecordID = file.id;
                            string RecordName = file.name;
                            ModifiedDate = file.ModifiedDate;
                            string Filename = file.Filename;
                            string ImageWidthText = file.Width;
                            string ImageHeightText = file.Height;
                            if (ModifiedDate <= DateTime.MinValue)
                                ModifiedDate = file.DateAdded;
                            string Description = file.Description;
                            string ImageAlt = file.AltText;
                            FileTypeID = file.FileTypeID;
                            int fileSize = file.FileSize;
                            AltSizeList = file.AltSizeList;
                            // 
                            string ImageSrc = cp.Site.FilePath + Strings.Replace(Filename, @"\", "/");
                            // 
                            int DotPosition = Strings.InStrRev(ImageSrc, ".");
                            if (DotPosition == 0) {
                                FileExtension = "";
                                FilenameNoExtension = "";
                            } else {
                                FileExtension = Strings.UCase(Strings.Mid(ImageSrc, DotPosition + 1));
                                FilenameNoExtension = Strings.Mid(ImageSrc, 1, DotPosition - 1);
                            }
                            // 
                            if (FileTypeID == 0) {
                                FileTypeID = getFileTypeID(cp, ImageSrc);
                                if (FileTypeID != 0)
                                    UpdateRecord = true;
                            }
                            // 
                            // if no name given, use the filename
                            // 
                            if (RecordName == "") {
                                if (ImageSrc == "")
                                    RecordName = "[no name]";
                                else {
                                    DotPosition = Strings.InStrRev(ImageSrc, "/");
                                    if (DotPosition == 0)
                                        RecordName = ImageSrc;
                                    else
                                        RecordName = Strings.Mid(ImageSrc, DotPosition + 1);
                                }
                            }
                            file.name = RecordName;

                            // 
                            string ResourceHref = "";
                            IconLink = "";
                            if (AllowFileAuthoring)
                                EditLink = adminUrl(cp) + "?cid=" + FileCID + "&id=" + ResourceRecordID + "&af=4" + "&aa=2&depth=1";
                            else
                                EditLink = "";
                            string ThumbNailSrc;
                            // 
                            // create thumbnail
                            // 
                            if (AllowThumbnails) {
                                ThumbNailSrc = ImageSrc;
                                if ((FilenameNoExtension != "") & (AltSizeList != "")) {
                                    string[] AltSizes = Strings.Split(AltSizeList, Constants.vbCrLf);
                                    int BestFitHeight = 9999999;
                                    string BestFitAltSize = "";
                                    for (Ptr = 0; Ptr <= Information.UBound(AltSizes); Ptr++) {
                                        // 
                                        // Find the smallest image larger then height 50
                                        // 
                                        string AltSize = Strings.Trim(AltSizes[Ptr]);
                                        if (AltSize != "") {
                                            Pos = Strings.InStr(AltSize, "x");
                                            if (Pos > 0) {
                                                int AltSizeHeight = cp.Utils.EncodeInteger(Strings.Mid(AltSize, Pos + 1));
                                                if (AltSizeHeight >= 50 & AltSizeHeight < BestFitHeight) {
                                                    BestFitHeight = AltSizeHeight;
                                                    BestFitAltSize = AltSize;
                                                }
                                            }
                                        }
                                    }
                                    if (BestFitAltSize != "")
                                        ThumbNailSrc = FilenameNoExtension + "-" + BestFitAltSize + "." + FileExtension;
                                }
                            }
                            // get file size
                            // 
                            // FileSize = 0
                            if (fileSize == 0) {
                                string Pathname = cp.Site.PhysicalFilePath + Strings.Replace(Filename, "/", @"\");
                                fileSize = GetFileSize(cp, Pathname);
                                if (fileSize != 0)
                                    UpdateRecord = true;
                            }
                            // 
                            // 
                            // 
                            if (UpdateRecord)
                                cp.Db.ExecuteSQL("update cclibraryFiles set FileTypeID=" + FileTypeID + ",filesize=" + fileSize + " where ID=" + ResourceRecordID);
                            // 
                            ImageSrc = kmaEncodeURL(cp, ImageSrc);
                            string IconOnClick = "";
                            FormDetails = FormDetails + GetFormRow_Files(cp, fileSize, IconLink, IconOnClick, RecordName, ImageSrc, ModifiedDate, RowCount, EditLink, Description, FileExtension, RecordName, ImageSrc, ImageAlt, ImageWidthText, ImageHeightText, ResourceRecordID, currentFolderID, AllowThumbnails, FileTypeFilter, ThumbNailSrc, SourceMode, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn);
                            RowCount = RowCount + 1;
                        }
                        // 
                        // ----- If nothing found, print no files found
                        if (RowCount == 0) {
                            FormDetails = FormDetails + "<tr class=\"listRow\"><td class=\"center\">" + IconSpacer + "</td><td class=\"left\" colspan=" + ColumnCnt - 1 + ">no folders or files were found</td></tr>";
                            RowCount = RowCount + 1;
                        }
                    }
                    // 
                    // Fill out the table to MinRows
                    // 
                    hint = "800";
                    for (RowCount = RowCount; RowCount <= iMinRows; RowCount++)
                        FormDetails = FormDetails + GetFormRow_Blank(cp, "", "", "", "", "", null/* TODO Change to default(_) if this is not a reference type */, RowCount, "", "", "BLANK", "", "", "", "", "", 0, currentFolderID, AllowEditColumn, AllowPlaceColumn, AllowSelectColumn, ColumnCnt);
                    // 
                    // Upload link
                    // 
                    if (AllowLocalFileAdd) {
                        // 
                        // Upload Form
                        // 
                        FormDetails = FormDetails + GetFormRow_Options(cp, currentFolderID, topFolderPath, ColumnCnt, IsContentManagerFiles, IsContentManagerFolders, currentFolderHasModifyAccess);
                        RowCount = RowCount + 1;
                    }
                    // 
                    // Bottom border
                    // 
                    // FormDetails = FormDetails & "<tr class=""border""><td class=""border"" Colspan=" & (ColumnCnt) & ">" & cp.Html.div("&nbsp;") & "</td></tr>"
                    FormDetails = FormDetails + "</table>";
                    // 
                    // Create the FormFolders
                    // 
                    string FormFolders = GetRLNav(cp, currentFolderID, topFolderPath, topFolderID);
                    FormFolders = "<div class=\"rlnav\">" + FormFolders + "</div>";
                    // FormFolders = Main.GetPanelRev(FormFolders)
                    // 
                    // Assemble the form
                    // 
                    hint = "900";
                    result = result + "<table border=0 cellpadding=0 cellspacing=0 width=\"100%\"><tr>";
                    if (!blockFolderNavigation) {
                        result = result + "<td class=\"nav ccPanel3DInput\">" + FormFolders + "<BR><img src=/ResourceLibrary/spacer.gif width=140 height=1></td>";
                        result = result + "<td class=\"navBorder ccPanel3D\"><img src=/ResourceLibrary/spacer.gif width=5 height=1></td>";
                    }
                    result = result + "<td class=\"content\">" + FormDetails + "</td>";
                    result = result + "</tr></Table>";
                    result = result + ButtonBar;
                    result = result + htmlHidden("RowCount", RowCount);
                    result = cp.Html.Form(result);
                }
                // 
                result = "<div class=\"ccLibrary\">" + result + "</div>";
            }
            // 
            // Help Link
            // 
            // result = Main.GetHelpLink(42, "Using the Resource Library", "The Resource Library is a convenient place to store reusable content, such as images and downloads. Objects in the Library can be placed on any page. The Library itself can be added to any page on your site.") & GetForm
            catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        // 
        //====================================================================================================
        // 
        private string GetFormRow_Folders(CPBaseClass cp, string ignore0, string IconLink, string IconOnClick, string Name, string NameLink, DateTime ModifiedDate, int RowCount, string EditLink, string Description, string FileType, string ResourceName, string ResourceLink, string ImageAlt, string ImageWidth, string ImageHeight, int RecordID, int FolderID, bool AllowEditColumn, bool AllowPlaceColumn) {
            string result = "";
            // 
            try {
                string RowClass;
                // 
                if ((RowCount % 2) == 0)
                    RowClass = "ccPanelRowOdd";
                else
                    RowClass = "ccPanelRowEven";
                // 

                string CellStart = "<td class=\"left ccAdminSmall " + RowClass + "\">";
                string CellStartRight = "<td class=\"right ccAdminSmall " + RowClass + "\">";
                string CellStart2 = "<td class=\"left ccAdminSmall " + RowClass + "\">";
                string CellStart5 = "<td class=\"left ccAdminSmall " + RowClass + "\">";
                string CellEnd = "</td>";
                string DateString;
                // 
                if (ModifiedDate <= DateTime.MinValue)
                    DateString = "&nbsp;";
                else
                    DateString = Strings.FormatDateTime(ModifiedDate, Constants.vbShortDate);
                // 
                result = result + "<tr class=\"row " + RowClass + "\">";
                result = result + CellStart + "&nbsp;" + CellEnd;
                if (AllowEditColumn)
                    result = result + CellStart + "&nbsp;" + CellEnd;
                if (AllowPlaceColumn)
                    result = result + CellStart + "&nbsp;" + CellEnd;
                result = result + CellStart + "<A href=\"?" + cp.Utils.EncodeUrl(IconLink) + "\">" + IconFolderOpen + "</A>" + CellEnd;
                result = result + CellStart + Name + CellEnd;
                result = result + CellStart + Description + CellEnd;
                result = result + CellStart + "&nbsp;" + CellEnd;
                result = result + CellStartRight + DateString + CellEnd;
                result = result + "</tr>";
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        // 
        //====================================================================================================
        // 
        private string GetFormRow_ChildFolders(CPBaseClass cp, string ignore0, string IconLink, string IconOnClick, string Name, string NameLink, DateTime ModifiedDate, int RowCount, string EditLink, string Description, string FileType, string ResourceName, string ResourceLink, string ImageAlt, string ImageWidth, string ImageHeight, int RecordID, int FolderID, bool AllowEditColumn, bool AllowPlaceColumn, bool AllowSelectColumn) {
            string result = "";
            // 
            try {
                // 
                string RowClass;
                // 
                if ((RowCount % 2) == 0)
                    RowClass = "ccPanelRowOdd";
                else
                    RowClass = "ccPanelRowEven";
                // 
                string CellStart = Constants.vbCrLf + "<td class=\"left ccAdminSmall\">";
                string CellStartCenter = Constants.vbCrLf + "<td class=\"center ccAdminSmall\">";
                string CellStartRight = Constants.vbCrLf + "<td class=\"right ccAdminSmall\">";
                string CellEnd = "</td>";
                string DateString;
                // 
                if (ModifiedDate <= DateTime.MinValue)
                    DateString = "&nbsp;";
                else
                    DateString = Strings.FormatDateTime(ModifiedDate, Constants.vbShortDate);
                if (Description == "")
                    Description = "&nbsp;";
                // 
                GetFormRow_ChildFolders = GetFormRow_ChildFolders + Constants.vbCrLf + "<tr class=\"listRow\" ID=\"Row" + RowCount + "\">";
                if (AllowSelectColumn)
                    GetFormRow_ChildFolders = GetFormRow_ChildFolders + CellStartCenter + "<input type=checkbox ID=Select" + RowCount + " name=Row" + RowCount + " value=1 onClick=\"RLRowClick(this.checked,'Row" + RowCount + "');\">" + htmlHidden("Row" + RowCount + "FolderID", FolderID) + CellEnd;
                if (AllowEditColumn) {
                    if (EditLink != "")
                        GetFormRow_ChildFolders = GetFormRow_ChildFolders + CellStartCenter + "<A href=\"" + EditLink + "\">" + IconFolderEdit + "</A>" + CellEnd;
                    else
                        GetFormRow_ChildFolders = GetFormRow_ChildFolders + CellStart + "&nbsp;" + CellEnd;
                }
                if (AllowPlaceColumn)
                    GetFormRow_ChildFolders = GetFormRow_ChildFolders + CellStartCenter + IconNoFile + CellEnd;
                GetFormRow_ChildFolders = GetFormRow_ChildFolders + CellStartCenter + "<A href=\"?" + IconLink + "\">" + IconFolderClosed + "</A>" + CellEnd;
                GetFormRow_ChildFolders = GetFormRow_ChildFolders + CellStart + "<A href=\"?" + IconLink + "\">" + Name + "</A>" + CellEnd;
                GetFormRow_ChildFolders = GetFormRow_ChildFolders + CellStart + Description + CellEnd;
                GetFormRow_ChildFolders = GetFormRow_ChildFolders + CellStartRight + "&nbsp;" + CellEnd;
                GetFormRow_ChildFolders = GetFormRow_ChildFolders + CellStartRight + DateString + CellEnd;
                GetFormRow_ChildFolders = GetFormRow_ChildFolders + "</tr>";
                result = GetFormRow_ChildFolders;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        // 
        //====================================================================================================
        // 
        private string GetFormRow_Files(CPBaseClass cp, int fileSize, string IconLink, string IconOnClick, string Name, string NameLink, DateTime ModifiedDate, int RowCount, string EditLink, string Description, string FilenameExt, string ResourceName, string ResourceLink, string ImageAlt, string ImageWidth, string ImageHeight, int RecordID, int FolderID, bool AllowThumbnails, string FileTypeFilter, string ThumbNailSrc, int SourceMode, bool AllowEditColumn, bool AllowPlaceColumn, bool AllowSelectColumn) {
            string result = "";
            // 
            try {
                string RowClass;
                // 
                if ((RowCount % 2) == 0)
                    RowClass = "ccPanelRowOdd";
                else
                    RowClass = "ccPanelRowEven";
                // 
                string CellStart = Constants.vbCrLf + "<td class=\"left ccAdminSmall\">";
                string CellStartCenter = Constants.vbCrLf + "<td class=\"center ccAdminSmall\">";
                string CellStartRight = Constants.vbCrLf + "<td class=\"right ccAdminSmall\">";
                string CellEnd = "</td>";
                string DateString;
                // 
                if (ModifiedDate <= DateTime.MinValue)
                    DateString = "&nbsp;";
                else
                    DateString = Strings.FormatDateTime(ModifiedDate, Constants.vbShortDate);
                // 
                // Determine Icons and actions
                // 
                bool AllowPlace;
                AllowPlace = false;
                string IconIMG;
                string IconFilename;
                bool IsImage;
                bool IsVideo;
                bool IsFlash;
                bool IsDownload;
                string Downloadfilename;
                string FileTypeName;
                string TestFileTYpe;
                bool FileTypeFound;
                string MediaIMG = "";
                if (iconFiles.Count <= 0)
                    IconIMG = IconImage;
                else {
                    TestFileTYpe = "," + Strings.UCase(Strings.Replace(FilenameExt, ".", "")) + ",";
                    foreach (FileType iconFile in iconFiles) {
                        int FileTypePtr;
                        if (Strings.InStr(1, "," + iconFiles[FileTypePtr].ExtensionList + ",", Strings.UCase(TestFileTYpe), Constants.vbTextCompare) != 0) {
                            {
                                var withBlock = iconFile;
                                FileTypeName = withBlock.Name;
                                IsImage = withBlock.IsImage;
                                IsVideo = withBlock.IsVideo;
                                IsFlash = withBlock.IsFlash;
                                bool IsMedia = IsImage | IsVideo | IsFlash;
                                // 
                                // 4/15/08 - if no filter, show everything
                                // 
                                // MediaIMG = IconNoFile

                                // If FileTypeFilter = "image" And IsImage Then
                                // MediaIMG = "<img src=""/ResourceLibrary/IconImagePlace2322.gif"" width=23 height=22 border=0 alt=""Place this image on the page"">"
                                // AllowPlace = True
                                // ElseIf FileTypeFilter = "media" And IsVideo Then
                                // MediaIMG = "<img src=""/ResourceLibrary/IconImagePlace2322.gif"" width=23 height=22 border=0 alt=""Place this video on the page"">"
                                // AllowPlace = True
                                // ElseIf FileTypeFilter = "flash" And IsVideo Then
                                // MediaIMG = "<img src=""/ResourceLibrary/IconImagePlace2322.gif"" width=23 height=22 border=0 alt=""Place this flash on the page"">"
                                // AllowPlace = True
                                // Else
                                // MediaIMG = IconNoFile
                                // AllowPlace = False
                                // End If
                                if (withBlock.MediaIconFilename != "")
                                    MediaIMG = "<img src=\"" + withBlock.MediaIconFilename + "\" width=23 height=22 border=0 alt=\"Place this flash on the page\">";
                                IsDownload = withBlock.IsDownload;
                                Downloadfilename = withBlock.DownloadIconFilename;
                                IconFilename = withBlock.IconFilename;
                                if (IconFilename == "")
                                    IconFilename = "/ResourceLibrary/IconDefault.gif";
                                IconIMG = "<img src=\"" + IconFilename + "\" border=\"0\" width=\"22\" height=\"23\" alt=\"" + withBlock.Name + "\">";
                            }
                            FileTypeFound = true;
                            break;
                        }
                    }
                }
                // 
                if (!FileTypeFound) {
                    if (FilenameExt == "PNG")
                        IsImage = true;
                    else if (FilenameExt == "JPG")
                        IsImage = true;
                    else if (FilenameExt == "GIF")
                        IsImage = true;
                    else
                        IsImage = false;
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
                    if (IsImage)
                        AllowPlace = true;
                } else if (FileTypeFilter == "media") {
                    if (IsVideo)
                        AllowPlace = true;
                } else if (FileTypeFilter == "flash") {
                    if (IsFlash)
                        AllowPlace = true;
                } else
                    // 
                    // no filter - place anything
                    // 
                    AllowPlace = true;
                if (AllowPlace & MediaIMG == "")
                    MediaIMG = "<img src=\"/ResourceLibrary/IconImagePlace2322.gif\" width=23 height=22 border=0 alt=\"Place this file on the page\">";
                // 
                // Output the row
                // 
                GetFormRow_Files = GetFormRow_Files + Constants.vbCrLf + "<tr class=\"listRow\" ID=\"Row" + RowCount + "\">";
                if (AllowSelectColumn)
                    GetFormRow_Files = GetFormRow_Files + CellStartCenter + "<input type=checkbox ID=Select" + RowCount + " name=Row" + RowCount + " value=1 onClick=\"RLRowClick(this.checked,'Row" + RowCount + "');\">" + htmlHidden("Row" + RowCount + "FileID", RecordID) + CellEnd;
                // 
                // ----- Edit Column
                // 
                if (AllowEditColumn) {
                    if (EditLink != "")
                        GetFormRow_Files = GetFormRow_Files + CellStartCenter + "<A href=\"" + EditLink + "\">" + IconFileEdit + "</A>" + CellEnd;
                    else
                        GetFormRow_Files = GetFormRow_Files + CellStart + "&nbsp;" + CellEnd;
                }
                // 
                // ----- Place Column
                // 
                if (!AllowPlaceColumn) {
                } else if ((!AllowPlace))
                    // 
                    // Can not select resources - display dot
                    // 
                    GetFormRow_Files = GetFormRow_Files + CellStartCenter + IconNoFile + CellEnd;
                else {
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
                        GetFormRow_Files = GetFormRow_Files + CellStartCenter + "<a href=\"#\" onClick=\"var e=window.opener.document.getElementById('" + SelectLinkObjectName + "');e.value='" + JSCopy + "'; window.close();\">" + ImageLink + "</A>" + CellEnd;
                    } else if (SourceMode == SourceModeFromDownloadRequest) {
                        // 
                        // return a simple download
                        // 
                        if (IsDownload) {
                            JSCopy = Downloadfilename;
                            JSCopy = Strings.Replace(JSCopy, @"\", @"\\");
                            JSCopy = kmaEncodeJavascript(cp, JSCopy);
                            ImageLink = "<img src=\"/ResourceLibrary/IconDownload2.gif\" border=\"0\" width=\"23\" height=\"22\" alt=\"Link to this resource\" title=\"Link to this resource\" valign=\"absmiddle\">";
                            GetFormRow_Files = GetFormRow_Files + CellStartCenter + "<a href=\"#\" onClick=\"window.opener.InsertDownload( '" + RecordID + "','" + SelectResourceEditorObjectName + "','" + JSCopy + "'); window.close();\">" + ImageLink + "</A>" + CellEnd;
                        } else
                            GetFormRow_Files = GetFormRow_Files + CellStartCenter + IconNoFile + CellEnd;
                    } else if (SourceMode == SourceModeFromLinkDialog) {
                        // 
                        // Return the file as a url to the editor dialog
                        // 
                        if (AllowPlace) {
                            JSCopy = kmaEncodeJavascript(cp, ResourceLink);
                            string JSClose = ""
                            + " if(navigator.appName.indexOf('Microsoft')!=-1) {window.returnValue='" + JSCopy + "'}"
                            + " else{window.opener.setAssetValue('" + JSCopy + "')}"
                            + " self.close();";
                            GetFormRow_Files = GetFormRow_Files + CellStartCenter + "<a href=\"#\" onClick=\"" + JSClose + "\" >" + MediaIMG + "</A>" + CellEnd;
                        } else
                            GetFormRow_Files = GetFormRow_Files + CellStartCenter + IconNoFile + CellEnd;
                    }
                }
                NameLink = cp.Utils.DecodeUrl(NameLink);

                GetFormRow_Files = GetFormRow_Files + CellStartCenter + IconIMG + CellEnd;
                GetFormRow_Files = GetFormRow_Files + CellStart + "<a href=\"" + NameLink + "\" target=\"_blank\">" + Name + "</A>" + CellEnd;
                // 
                if (Description == "")
                    Description = "&nbsp;";
                if (AllowThumbnails & IsImage)
                    // If AllowThumbnails And (UCase(FileTypeName) = "IMAGE") Then
                    GetFormRow_Files = GetFormRow_Files
            + CellStart
            + "<a href=\"" + NameLink + "\" target=\"_blank\">"
            + "<img src=\"" + ThumbNailSrc + "\" height=\"50\"  vspace=\"0\" hspace=\"10\" style=\"vertical-align:middle;border:0;\">"
            + "</a>"
            + Description
            + CellEnd;
                else
                    GetFormRow_Files = GetFormRow_Files
        + CellStart
        + Description
        + CellEnd;
                // 
                if (fileSize > 10000)
                    GetFormRow_Files = GetFormRow_Files + CellStartRight + Conversion.Int(fileSize / (double)1024) + "&nbsp;KB&nbsp;" + CellEnd;
                else
                    GetFormRow_Files = GetFormRow_Files + CellStartRight + fileSize + "&nbsp;" + CellEnd;
                // 
                GetFormRow_Files = GetFormRow_Files + CellStartRight + DateString + CellEnd;
                GetFormRow_Files = GetFormRow_Files + "</tr>";
                result = GetFormRow_Files;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        // 
        //====================================================================================================
        // 
        private string GetFormRow_Blank(CPBaseClass cp, string ignore0, string IconLink, string IconOnClick, string Name, string NameLink, DateTime ModifiedDate, int RowCount, string EditLink, string Description, string FileType, string ResourceName, string ResourceLink, string ImageAlt, string ImageWidth, string ImageHeight, int RecordID, int FolderID, bool AllowEditColumn, bool AllowPlaceColumn, bool AllowSelectColumn, int ColumnCnt) {

            // 
            GetFormRow_Blank = Constants.vbCrLf + Constants.vbTab + "<tr class=\"listRow\"><td class=\"left\"><img height=\"23\" width=\"22\" src=\"/wwwroot/ResourceLibrary/spacer.gif\"></td><td class=\"left\" colspan=\"" + ColumnCnt - 1 + "\">&nbsp;</td></tr>";
        }
        // 
        //====================================================================================================
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
                    // If IsContentManagerFolders Then
                    FolderCell = ""
                    + "<table id=\"AddFolderTable\" border=\"0\" cellpadding=\"10\" cellspacing=\"1\" width=\"100%\">"
                    + "<tr>";
                    FolderCell = FolderCell
                    + "<td class=\"left\" align=\"left\" colspan=2>" + kmaAddSpan("Add Folder&nbsp;", "ccAdminSmall") + "</td>"
                    + "<td class=\"left\" Width=\"99%\" align=\"left\">" + kmaAddSpan("Description&nbsp;", "ccAdminSmall") + "</td>"
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
                    + "<td class=\"left\" align=\"left\" colspan=2>" + kmaAddSpan("Add Files&nbsp;", "ccAdminSmall") + "</td>"
                    + "<td class=\"left\" Width=\"100\" align=\"left\">" + kmaAddSpan("Name&nbsp;", "ccAdminSmall") + "</td>"
                    + "<td class=\"left\" Width=\"100\" align=\"left\">" + kmaAddSpan("Description&nbsp;", "ccAdminSmall") + "</td>"
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
                GetFormRow_Options = ""
                    + "<div  style=\"margin-left:10px;\">" + cp.Html.CheckBox("AllowThumbnails", cp.User.GetBoolean("LibraryAllowthumbnails", "0")) + "&nbsp;Display Thumbnails";
                if (cp.User.IsAdmin | hasModifyAccess) {
                    // 
                    string moveSelect = GetMoveFolderPathSelect(cp, FolderID, topFolderPath);
                    if (moveSelect != "")
                        GetFormRow_Options += "<BR>" + cp.Html.CheckBox("Move", false) + "&nbsp;Move selected files to " + moveSelect;
                    if (FolderCell != "")
                        GetFormRow_Options += "<hr>" + cp.Html.div(FolderCell);
                    if (FileCell != "")
                        GetFormRow_Options += "<hr>" + cp.Html.div(FileCell);
                }
                if (GetFormRow_Options != "") {
                    GetFormRow_Options = cp.Html.div(GetFormRow_Options);
                    GetFormRow_Options = "<tr><td class=\"bg-light left\" colspan=" + (ColumnCnt) + ">" + GetFormRow_Options + "</td></tr>";
                }
                // 
                result = GetFormRow_Options;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        // 
        //====================================================================================================
        // 
        private string GetForm_HeaderCell(CPBaseClass cp, string Align, string Width, string Copy) {
            string Style = ""
                   + "padding: 3px;"
                   + "font-size:10px;";
            string result = "<td WIDTH=\"" + Width + "\" ALIGN=\"" + Align + "\" class=ccAdminListCaption style=\"" + Style + "\">"
                    + Copy
                    + "</td>";
            return result;
        }
        // 
        //====================================================================================================
        // 
        private bool IsInFolder(CPBaseClass cp, int topFolderID, int FolderID, string ParentPath = "") {
            try {
                if ((FolderID == 0))
                    return false;
                if ((topFolderID == 0))
                    return true;
                if ((Strings.InStr(1, "," + ParentPath + ",", "," + System.Convert.ToString(FolderID) + ",") != 0))
                    return false;
                ParentPath += "," + System.Convert.ToString(FolderID);
                var folder = LibraryFolderModel.create(cp, FolderID);
                int ParentID;
                if ((folder != null))
                    ParentID = folder.ParentID;
                if (ParentID == 0)
                    return false;
                else if (ParentID == topFolderID)
                    return true;
                else
                    return IsInFolder(cp, topFolderID, ParentID, ParentPath);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return IsInFolder;
        }
        // 
        //====================================================================================================
        // 
        private string GetParentFoldersLink(CPBaseClass cp, string topFolderPath, int topFolderID, int currentFolderID, int FolderID, string RefreshQS, string ChildIDList) {
            string result = "";
            try {
                string folderName = "";
                if ((FolderID == 0) | (FolderID == topFolderID)) {
                    // 
                    // Root folder
                    folderName = topFolderPath;
                    if (folderName == "")
                        folderName = "Root";
                    if (currentFolderID == FolderID)
                        result = "Folder <B>" + folderName + "</B>";
                    else
                        result = "Folder <a href=?" + RefreshQS + "&FolderID=0>" + folderName + "</a>";
                } else {
                    LibraryFolderModel LibraryFolder = LibraryFolderModel.create(cp, "ID=" + FolderID);
                    int ParentID;
                    bool RecordFound;
                    if (!(LibraryFolder == null)) {
                        RecordFound = true;
                        ParentID = LibraryFolder.ParentID;
                        folderName = LibraryFolder.name;
                    }
                    string FolderLink;
                    // 
                    if (currentFolderID == FolderID)
                        FolderLink = "<B>" + folderName + "</B>";
                    else
                        FolderLink = "<a href=?" + RefreshQS + "&FolderID=" + FolderID + ">" + folderName + "</a>";
                    if ((!RecordFound) | (FolderID == topFolderID)) {
                        // 
                        // call this the top of the tree
                        if (folderName == "")
                            folderName = "Root";
                        result = "Folder " + FolderLink;
                    } else if (Strings.InStr(1, ChildIDList + ",", "," + FolderID + ",") != 0)
                        // 
                        // circular reference - end it here
                        result = "Folder (Circular Reference) > " + FolderLink;
                    else if (currentFolderID == ParentID)
                        // 
                        // circular reference - end it here
                        result = "Folder " + FolderLink;
                    else
                        result = GetParentFoldersLink(cp, topFolderPath, topFolderID, currentFolderID, ParentID, RefreshQS, ChildIDList + "," + FolderID) + @"\" + FolderLink;
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        // 
        //====================================================================================================
        // 
        private string GetFolderPathSelect(CPBaseClass cp, int topFolderPathID, string topFolderPath, bool RequireModifyAccess) {
            string result = "";
            try {
                string pathRemoveString = "";
                string pathCaption;
                // 
                // GetFolderPathSelect = FolderSelect
                if (GetFolderPathSelect == "") {
                    // 
                    // 
                    // create full paths, set .hasViewAccess
                    // 
                    int optionCnt = 0;
                    if (topFolderPath != "") {
                        pathRemoveString = @"root\";
                        int Pos = Strings.InStrRev(topFolderPath, @"\");
                        if (Pos > 0)
                            pathRemoveString = pathRemoveString + Strings.Mid(topFolderPath, 1, Pos - 1);
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
                        if (topFolderPathID == 0)
                            // 
                            // if root is current folder, mark it selected
                            // 
                            GetFolderPathSelect = GetFolderPathSelect + "<option value=0 selected>Root</option>";
                        else
                            GetFolderPathSelect = GetFolderPathSelect + "<option value=0>Root</option>";
                    }
                    int Ptr = FolderPathIndex.GetFirstPointer;
                    while ((Ptr >= 0)) {
                        if (folders[Ptr].hasViewAccess & ((!RequireModifyAccess) | folders[Ptr].hasModifyAccess)) {
                            int PtrFolderID = folders[Ptr].FolderID;
                            pathCaption = Strings.Replace(folders[Ptr].FullPath, pathRemoveString, "", Compare: Constants.vbTextCompare);

                            if (PtrFolderID == topFolderPathID)
                                GetFolderPathSelect = GetFolderPathSelect + "<option value=" + PtrFolderID + " selected>" + pathCaption + "</option>";
                            else
                                GetFolderPathSelect = GetFolderPathSelect + "<option value=" + PtrFolderID + ">" + pathCaption + "</option>";
                            optionCnt = optionCnt + 1;
                        }
                        Ptr = FolderPathIndex.GetNextPointer;
                    }
                    // 
                    // Create Select
                    // 
                    if (optionCnt <= 1)
                        // 
                        // If only one folder, (the current one), return nothing
                        // 
                        GetFolderPathSelect = "";
                    else
                        // If GetFolderPathSelect <> "" Then
                        GetFolderPathSelect = "<select name=FieldName size=1 onChange>" + GetFolderPathSelect + "</select>";
                    FolderSelect = GetFolderPathSelect;
                }
                // 
                result = FolderSelect;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
            return;
        }
        // 
        //====================================================================================================
        // 
        private string GetFolderPath(CPBaseClass cp, int targetPtr, string ChildIDList) {
            string result = "";
            try {
                // 
                int ParentPtr;
                int ParentID;
                int FolderID;

                // 
                result = folders[targetPtr].Name;
                ParentID = folders[targetPtr].parentFolderID;
                FolderID = folders[targetPtr].FolderID;
                if (ParentID == 0)
                    // 
                    // At the Root page
                    // 
                    result = @"Root\" + result;
                else if ((FolderID == ParentID) | (Strings.InStr(1, "," + ChildIDList + ",", "," + ParentID + ",") != 0)) {
                } else
                    for (ParentPtr = 0; ParentPtr <= Information.UBound(folders); ParentPtr++) {
                        // 
                        // todo Folder(parentPtr) throws a null ref this needs to be resolved
                        if ((folders[ParentPtr] == null))
                            cp.Utils.AppendLogFile("getfolderPath=6b ******** parentPtr [" + ParentPtr + "]");
                        else if (folders[ParentPtr].FolderID == ParentID) {
                            result = GetFolderPath(cp, ParentPtr, ChildIDList + "," + ParentID) + @"\" + result;
                            // GetFolderPath = GetFolderPath(ParentPtr, ChildIDList & "," & ParentID) & " > " & GetFolderPath
                            break;
                        }
                    }
            }
            // 

            catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
            return;
        }
        // 
        //====================================================================================================
        // 
        private string GetJumpFolderPathSelect(CPBaseClass cp, int FolderID, string topFolderPath) {
            string result = "";
            try {
                // 
                result = GetFolderPathSelect(cp, FolderID, topFolderPath, false);
                if (result != "") {
                    result = Strings.Replace(result, "FieldName", "JumpFolderID");
                    result = Strings.Replace(result, "onChange", "onChange=\"QJump(this);\" ");
                    result = Replace(result, "value=", "value=?" + cp.Doc.RefreshQueryString + "&FolderID=");
                    result = "<script language=JavaScript1.2>function QJump(e){var l=e.value;if(l!=''){window.name='RL';window.location.assign(l);}}</script>" + result;
                }
            }
            // 
            catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        // 
        //====================================================================================================
        // 
        private string GetMoveFolderPathSelect(CPBaseClass cp, int FolderID, string topFolderPath) {
            try {
                // 
                GetMoveFolderPathSelect = GetFolderPathSelect(cp, FolderID, topFolderPath, true);
                GetMoveFolderPathSelect = Strings.Replace(GetMoveFolderPathSelect, "FieldName", "MoveFolderID");
                GetMoveFolderPathSelect = Strings.Replace(GetMoveFolderPathSelect, "onChange", "onChange=\"var e=getElementById('Move');if(e){e.checked=true};\" ");
            }
            // 
            catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
        }
        // 
        //====================================================================================================
        // 
        private string GetRLNav(CPBaseClass cp, int currentFolderID, string topFolderPath, int topFolderID) {
            bool IsAuthoring;
            try {
                // 
                IsAuthoring = false;
                string BakeName = "RLNav";
                if (!IsAuthoring) {
                }
                if (GetRLNav == "") {
                    string LinkBase = cp.Doc.RefreshQueryString;
                    LinkBase = cp.Utils.ModifyQueryString(LinkBase, "FolderID", "0");

                    // 
                    // 

                    menuTreeClass Tree = new menuTreeClass(cp);
                    if (topFolderID == 0)
                        Tree.AddEntry(System.Convert.ToString(0), System.Convert.ToString(-1), "", "", "?" + LinkBase, "Root");
                    if (folderCnt > 0) {
                        int Ptr;
                        for (Ptr = 0; Ptr <= folderCnt - 1; Ptr++) {
                            int Id = folders[Ptr].FolderID;
                            if (folders[Ptr].hasViewAccess) {
                                // If hasModifyAccessByFolder(Id, topFolderPath) Then
                                int ParentID = folders[Ptr].parentFolderID;
                                string Caption = Strings.Replace(folders[Ptr].Name, " ", "&nbsp;");
                                string Link = "?" + cp.Utils.ModifyQueryString(LinkBase, "FolderID", System.Convert.ToString(Id));
                                Tree.AddEntry(System.Convert.ToString(Id), System.Convert.ToString(ParentID), "", "", Link, Caption);
                            }
                        }
                    }
                    GetRLNav = Tree.GetTree(System.Convert.ToString(topFolderID), System.Convert.ToString(currentFolderID));
                }
                // '
                // ' Get topFolderPath
                // '
                // If topFolderPath = "" Then
                // topFolderPath = "Root"
                // Else
                // topFolderPath = topFolderPath
                // End If
                // 
                // open the current node
                // 

                // Call main.AddOnLoadJavascript("convertTrees(); expandToItem('tree0','" & currentFolderID & "');")
                cp.Doc.AddOnLoadJavascript("convertTrees(); expandToItem('tree0','" + currentFolderID + "');");
            }
            // Link = "?" & LinkBase
            // Link = "<div style=""position:relative;left:-10;margin-bottom:3px;""><a href=" & Link & " style=""text-decoration:none ! important;"">" & topFolderPath & "</a></div>"
            // GetRLNav = Replace(GetRLNav, "<LI ", Link & "<LI ", 1, 1, vbTextCompare)
            // 'If CurrentFolderID <> 0 Then
            // GetRLNav = GetRLNav & "<script type=""text/javascript"">convertTrees(); expandToItem('tree0','" & CurrentFolderID & "');</script>"
            // 'End If
            catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
        }
        // 
        //====================================================================================================
        // 
        private bool AllowFolderAccess(CPBaseClass cp, int FolderID, int ParentID) {
            bool result = false;
            try {
                int GrandParentID;
                if (FolderID == 0 || cp.User.IsAdmin) { return true; };
                AllowFolderAccess = true;
                List<Contensive.Models.Db.LibraryFolderModel> LibraryFolderModelList = Models.LibraryFolderModel.AllowFolderAccess(cp, FolderID, ParentID);
                // 
                if (!AllowFolderAccess & (ParentID != 0)) {
                    LibraryFolderModel LibraryFolder = LibraryFolderModel.create(cp, ParentID);
                    if (!(LibraryFolder == null))
                        GrandParentID = LibraryFolder.ParentID;
                    AllowFolderAccess = AllowFolderAccess(cp, ParentID, GrandParentID);
                }
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return false;
            }
        }
        // 
        //====================================================================================================
        // 
        private bool hasModifyAccessByFolder(CPBaseClass cp, int FolderID, string topFolderPath) {
            try {
                // 
                int Ptr;
                // 
                if (FolderID == 86)
                    FolderID = FolderID;

                // 
                if (cp.User.IsAdmin)
                    // 
                    // 
                    // 
                    hasModifyAccessByFolder = true;
                else {
                    // 
                    // Need to check permissions
                    // 
                    LoadFolders_returnTopFolderId(cp, topFolderPath);
                    if (FolderID == 0)
                        hasModifyAccessByFolder = true;
                    else {
                        Ptr = FolderIdIndex.getPtr(System.Convert.ToString(FolderID));
                        if (Ptr >= 0)
                            hasModifyAccessByFolder = folders[Ptr].hasModifyAccess;
                    }
                }
            }
            // 
            catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
        }
        // 
        // ====================================================================================================
        // 
        private int LoadFolders_returnTopFolderId(CPBaseClass cp, string topFolderPath) {
            int topFolderID;
            try {
                FolderIdIndex = new FastIndexClass();
                FolderNameIndex = new FastIndexClass();
                FolderPathIndex = new FastIndexClass();
                // 
                // Load the folders storage
                List<Models.LibraryFolderModel> foldersList = Models.LibraryFolderModel.createList(cp, "");
                folderCnt = 0;
                if ((foldersList.Count > 0)) {
                    // 
                    // Store folders and setup folder index
                    // 
                    folderCnt = foldersList.Count;
                    folders = new FolderType[foldersList.Count - 1 + 1];
                    int Ptr = 0;

                    foreach (var folder in foldersList) {
                        folders[Ptr] = new FolderType();
                        if (true) {
                            if (true) {
                                FolderIdIndex.SetPointer(System.Convert.ToString(folder.id), Ptr);
                                FolderNameIndex.SetPointer(folder.name, Ptr);
                                {
                                    var withBlock = folders[Ptr];
                                    withBlock.FolderID = folder.id;
                                    withBlock.parentFolderID = folder.ParentID;
                                    withBlock.Name = folder.name;
                                    withBlock.hasModifyAccess = true;
                                    withBlock.modifyAccessIsValid = true;
                                    withBlock.hasViewAccess = true;
                                }
                                // 
                                // FullPath, propigate modifyAccess from parent to folder , ViewAccess
                                // 
                                {
                                    var withBlock = folders[Ptr];
                                    // 
                                    // determine modify access
                                    // 
                                    if ((!withBlock.modifyAccessIsValid)) {
                                        withBlock.hasModifyAccess = LoadFolders_GetModifyAccess(cp, withBlock.parentFolderID);
                                        withBlock.modifyAccessIsValid = true;
                                    }
                                    // 
                                    string testFullPath = GetFolderPath(cp, Ptr, "");
                                    folders[Ptr].FullPath = testFullPath;
                                    // End If
                                    FolderPathIndex.SetPointer(testFullPath, Ptr);
                                    // 
                                    if (Strings.InStr(1, testFullPath, @"root\" + topFolderPath, Constants.vbTextCompare) == 1)
                                        // 
                                        withBlock.hasViewAccess = true;
                                }
                                // 
                                // 
                                topFolderID = 0;
                                if (topFolderPath != "") {
                                    string[] targetFolders = Strings.Split(topFolderPath, @"\");
                                    int targetFolderCnt = Information.UBound(targetFolders) + 1;
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
                                                } else
                                                    // 
                                                    // good match, this as the parent and find the next
                                                    // 
                                                    break;
                                                testFolderPtr = FolderNameIndex.GetNextPointerMatch(targetFolderName);
                                            }
                                            if (testFolderPtr >= 0)
                                                targetFolderId = folders[testFolderPtr].FolderID;
                                            else
                                                // 
                                                // folder not found, create it with the parent
                                                // 
                                                // cS = main.InsertCSRecord("Library Folders")

                                                // If main.IsCSOK(cS) Then
                                                if (!(folder == null)) {
                                                targetFolderId = folder.id;
                                                folder.name = targetFolderName;
                                                folder.ParentID = targetParentFolderID;
                                                // Call main.SetCS(cS, "name", targetFolderName)
                                                // Call main.SetCS(cS, "parentid", targetParentFolderID)
                                                folder.save(cp);
                                            }
                                            if (Ptr == (targetFolderCnt - 1))
                                                topFolderID = targetFolderId;
                                        }
                                    }
                                }
                            }
                        }
                        Ptr += 1;
                    }
                }
                LoadFolders_returnTopFolderId = topFolderID;
                // 
                topFolderID = LoadFolders_returnTopFolderId;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return topFolderID;
        }
        // 
        //====================================================================================================
        // 
        private int loadFolders_getFolderID(CPBaseClass cp, string[] targetArray, int targetArrayPtr) {

            // 
            int cachePtr;
            int cacheFolderID;
            int cacheParentFolderID;
            string targetFolderName;
            int targetFolderParentId;
            // 
            loadFolders_getFolderID = 0;
            targetFolderName = targetArray[targetArrayPtr];
            cachePtr = FolderNameIndex.getPtr(targetFolderName);
            while (cachePtr >= 0) {
                cacheFolderID = folders[cachePtr].FolderID;
                if (targetArrayPtr == 0) {
                    // 
                    // this was the top-most folder, return the non-zero cache id
                    // 
                    if (folders[cachePtr].parentFolderID != 0) {
                    } else {
                        // 
                        // top of target path matches records (parentid=0)
                        // 
                        loadFolders_getFolderID = cacheFolderID;
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
                    } else
                        // 
                        // parent folder found, check that its target folder matches the parent id of this folder
                        // 
                        if (targetFolderParentId == folders[cachePtr].parentFolderID) {
                        // 
                        // this folder is correct, return with it's ID
                        // 
                        loadFolders_getFolderID = cacheFolderID;
                        break;
                    } else {
                    }
                }
                cachePtr = FolderNameIndex.GetNextPointerMatch(targetFolderName);
            }
        }
        // 
        //====================================================================================================
        // 
        private bool LoadFolders_GetModifyAccess(CPBaseClass cp, int FolderID) {

            // 
            int Ptr;
            // 
            Ptr = FolderIdIndex.getPtr(System.Convert.ToString(FolderID));
            if (Ptr >= 0) {
                if (folders[Ptr].modifyAccessIsValid)
                    // 
                    // 
                    // 
                    LoadFolders_GetModifyAccess = folders[Ptr].hasModifyAccess;
                else if (folders[Ptr].parentFolderID == 0) {
                    // 
                    // Parent is root, this folder does not have access
                    // 
                    LoadFolders_GetModifyAccess = false;
                    folders[Ptr].hasModifyAccess = LoadFolders_GetModifyAccess;
                    folders[Ptr].modifyAccessIsValid = true;
                } else {
                    // 
                    // Parent is not root
                    // 
                    LoadFolders_GetModifyAccess = LoadFolders_GetModifyAccess(cp, folders[Ptr].parentFolderID);
                    folders[Ptr].hasModifyAccess = LoadFolders_GetModifyAccess;
                    folders[Ptr].modifyAccessIsValid = true;
                }
            }
        }
        // 
        //====================================================================================================
        // 
        private void Class_Initialize() {
            iMinRows = 10;
        }
        // 
        //====================================================================================================
        // 
        private void Class_Terminate() {
        }
        // 
        //====================================================================================================
        // 
        private int GetFileSize(CPBaseClass cp, string VirtualFilePathPage) {
            string hint;
            int result;
            try {
                // 
                hint = "1";
                // 
                hint = "2";
                VirtualFilePathPage = Strings.Replace(VirtualFilePathPage, "/", @"\");
                int SlashPosition = Strings.InStrRev(VirtualFilePathPage, @"\");
                string Filename = "";
                string Pathname = "";
                if (SlashPosition != 0) {
                    Filename = Strings.LCase(Strings.Mid(VirtualFilePathPage, SlashPosition + 1));
                    Pathname = Strings.Mid(VirtualFilePathPage, 1, SlashPosition - 1);
                }
                // 
                string FileDescriptor;
                FileDescriptor = cp.File.fileList(Pathname);
                hint = "3";
                if (FileDescriptor == "") {
                } else {
                    hint = "4";
                    string[] FileSplit2 = Strings.Split(FileDescriptor, Constants.vbCrLf);
                    // Call AppendLogFile("GetFileSize, FileDescriptor lines=" & UBound(FileSplit2))
                    hint = "5";
                    int Ptr;
                    for (Ptr = 0; Ptr <= Information.UBound(FileSplit2); Ptr++) {
                        string[] FileParts = Strings.Split(FileSplit2[Ptr], Constants.vbTab);
                        if (Information.UBound(FileParts) <= 5) {
                        } else if (Strings.LCase(FileParts[0]) == Filename) {
                            GetFileSize = cp.Utils.EncodeInteger(FileParts[5]);
                            // Call AppendLogFile("GetFileSize, match on " & FileParts(0))
                            break;
                        }
                        result = GetFileSize;
                    }
                    hint = "6";
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return GetFileSize;
        }
        // 
        //====================================================================================================
        // 
        private int getFileTypeID(CPBaseClass cp, string Filename) {

            // 
            string[] FileNameSplit;
            string FileExtension;
            int CSType;
            int DefaultFileTypeID;
            int cnt;
            int Ptr;
            string hint;
            // 
            FileNameSplit = Strings.Split(Filename, ".");
            FileExtension = FileNameSplit[Information.UBound(FileNameSplit)];
            // 
            // try to read if from IconFiles
            // 
            hint = "1";
            GetFileTypeID = DefaultFileTypeID;
            foreach (FileType iconFile in iconFiles) {
                hint = "2";
                if (Strings.InStr(1, "," + iconFile.ExtensionList + ",", "," + FileExtension + ",", Constants.vbTextCompare) != 0) {
                    hint = "3";
                    GetFileTypeID = iconFile.FileTypeID;
                    break;
                    if (Strings.LCase(iconFile.Name) == "default") {
                        hint = "4";
                        DefaultFileTypeID = iconFile.FileTypeID;
                    }
                }
            }
            // cnt = IconFileCnt
            // If cnt > 0 Then
            // For Ptr = 0 To cnt - 1
            // Next
            // End If
            hint = "5";
        }
    }
}
