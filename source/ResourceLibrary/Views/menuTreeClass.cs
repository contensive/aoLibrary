
using System;
using Contensive.BaseClasses;
using Contensive.Addons.ResourceLibrary.Controllers;
using static Contensive.Addons.ResourceLibrary.Controllers.genericController;

namespace Contensive.Addons.ResourceLibrary {
    public class menuTreeClass {
        private CPBaseClass cp;

        internal menuTreeClass(CPBaseClass cp) {
            this.cp = cp;
            EntryIndexName = new FastIndexClass();
            MenuFlyoutNamePrefix = $"id{(int)Math.Floor(9999 * new Random().NextDouble())}";
        }

        const int MenuStyleRollOverFlyout = 1;
        const int MenuStyleTree = 2;
        const int MenuStyleTreeList = 3;
        const int MenuStyleFlyoutDown = 4;
        const int MenuStyleFlyoutRight = 5;
        const int MenuStyleFlyoutUp = 6;
        const int MenuStyleFlyoutLeft = 7;
        const int MenuStyleHoverDown = 8;
        const int MenuStyleHoverRight = 9;
        const int MenuStyleHoverUp = 10;
        const int MenuStyleHoverLeft = 11;

        public class MenuEntryType {
            public string Caption;
            public string Name;
            public string ParentName;
            public string Link;
            public string Image;
            public string ImageOver;
            public string ImageOpen;
            public bool NewWindow;
            public string OnClick;
        }

        public string iMenuFilePath;
        public int iEntryCount;
        public int iEntrySize;
        public MenuEntryType[] iEntry;
        public int iTreeCount;
        public string iMenuCloseString;
        public string UsedEntries;
        public FastIndexClass EntryIndexName;
        public FastIndexClass EntryIndexID;
        public string MenuFlyoutNamePrefix;
        public string MenuFlyoutIcon_Local;
        public const bool newmode = true;

        public void AddEntry(string EntryName, string ParentiEntryName, string ImageLink, string ImageOverLink, string Link, string Caption, string OnClickJavascript = "", string Ignore1 = "", string ImageOpenLink = "", bool NewWindow = false) {
            int MenuEntrySize;
            string iEntryName;
            string UcaseEntryName;
            bool iNewWindow;
            iEntryName = KmaEncodeMissingText(EntryName, "").Replace(",", " ");
            UcaseEntryName = iEntryName.ToUpper();
            if ((iEntryName != "") && ($"{UsedEntries},".IndexOf($",{UcaseEntryName},", StringComparison.Ordinal) + 1) == 0) {
                UsedEntries = $"{UsedEntries},{UcaseEntryName}";
                if (iEntryCount >= iEntrySize) {
                    iEntrySize = iEntrySize + 10;
                    Array.Resize(ref iEntry, iEntrySize + 1);
                }
                iEntry[iEntryCount] = new MenuEntryType();
                iEntry[iEntryCount].Link = KmaEncodeMissingText(Link, "");
                iEntry[iEntryCount].Image = KmaEncodeMissingText(ImageLink, "");
                iEntry[iEntryCount].OnClick = KmaEncodeMissingText(OnClickJavascript, "");
                if (iEntry[iEntryCount].Image == "") {
                    iEntry[iEntryCount].Caption = KmaEncodeMissingText(Caption, iEntryName);
                } else {
                    iEntry[iEntryCount].Caption = KmaEncodeMissingText(Caption, "");
                }
                iEntry[iEntryCount].Name = UcaseEntryName;
                iEntry[iEntryCount].ParentName = KmaEncodeMissingText(ParentiEntryName, "").ToUpper();
                iEntry[iEntryCount].ImageOver = KmaEncodeMissingText(ImageOverLink, "");
                iEntry[iEntryCount].ImageOpen = KmaEncodeMissingText(ImageOpenLink, "");
                iEntry[iEntryCount].NewWindow = KmaEncodeMissingBoolean(NewWindow, false);
                EntryIndexName.SetPointer(UcaseEntryName, iEntryCount);
                iEntryCount = iEntryCount + 1;
            }
        }

        public string GetMenu(string MenuName, string StyleSheetPrefix = "") {
            return GetTree(MenuName, "", KmaEncodeMissingText(StyleSheetPrefix, "ccTree"));
        }

        private string GetMenuTreeBranch(string ParentName, string JSObject, string UsedEntries) {
            string result = "";
            int EntryPointer;
            string iUsedEntries;
            string JSChildObject;
            int SubMenuCount;
            iUsedEntries = UsedEntries;
            SubMenuCount = 0;
            for (EntryPointer = 0; EntryPointer <= iEntryCount - 1; EntryPointer++) {
                if (iEntry[EntryPointer].ParentName == ParentName) {
                    if (($"{iUsedEntries},".IndexOf($",{EntryPointer},", StringComparison.Ordinal) + 1) == 0) {
                        JSChildObject = $"{JSObject}.s[{SubMenuCount}]";
                        iUsedEntries = $"{iUsedEntries},{EntryPointer}";
                        result = result
                            + $"{JSChildObject} = new so(0,'{iEntry[EntryPointer].Caption}','{iEntry[EntryPointer].Link}','_blank',''); \r\n"
                            + GetMenuTreeBranch(iEntry[EntryPointer].Name, JSChildObject, iUsedEntries);
                        SubMenuCount = SubMenuCount + 1;
                    }
                }
            }
            return result;
        }

        private string GetMenuTreeList(string MenuName, string OpenNodesList) {
            string result = "";
            int EntryPointer;
            string UcaseMenuName;
            if (iEntryCount > 0) {
                UcaseMenuName = MenuName.ToUpper();
                EntryPointer = EntryIndexName.GetPointer(UcaseMenuName);
                result = GetMenuTreeListBranch2(EntryPointer, "", OpenNodesList);
                return result;
            }
            return result;
        }

        private string GetMenuTreeListBranch2(int NodePointer, string UsedEntriesList, string OpenNodesList) {
            string result = "";
            string Link;
            int EntryPointer;
            string UcaseNodeName;
            string Image;
            string Caption;
            if (iEntryCount > 0) {
                if (($",{NodePointer},".IndexOf($",{UsedEntriesList},", StringComparison.Ordinal) + 1) == 0) {
                    result = result + "<ul Style=\"list-style-type: none; margin-left: 20px\">";
                    Caption = iEntry[NodePointer].Caption;
                    Link = kmaEncodeHTML(cp, iEntry[NodePointer].Link);
                    if (Link != "") {
                        Caption = $"<A TARGET=\"_blank\" HREF=\"{Link}\">{Caption}</A>";
                    }
                    if (($",{OpenNodesList},".IndexOf($",{NodePointer},", StringComparison.Ordinal) + 1) == 0) {
                        Image = iEntry[NodePointer].Image;
                        result = result + $"<li><A HREF=\"?OpenNodesList={OpenNodesList}&OpenNode={NodePointer}\"><IMG SRC=\"{Image}\" HEIGHT=\"18\" WIDTH=\"18\" BORDER=0 ALT=\"Open Folder\" /></A>&nbsp;{Caption}</li>";
                    } else {
                        Image = iEntry[NodePointer].ImageOpen;
                        if (Image == "") {
                            Image = iEntry[NodePointer].Image;
                        }
                        result = result
                            + "<li>"
                            + $"<A HREF=\"?OpenNodesList={OpenNodesList}&CloseNode={NodePointer}\">"
                            + $"<IMG SRC=\"{Image}\" HEIGHT=\"18\" WIDTH=\"18\" BORDER=0 ALT=\"Close Folder\" />"
                            + $"</A>&nbsp;{Caption}</li>";
                        UcaseNodeName = iEntry[NodePointer].Name.ToUpper();
                        for (EntryPointer = 0; EntryPointer <= iEntryCount - 1; EntryPointer++) {
                            if (iEntry[EntryPointer].ParentName == UcaseNodeName) {
                                result = result + GetMenuTreeListBranch2(EntryPointer, $"{UsedEntriesList},{NodePointer}", OpenNodesList);
                            }
                        }
                    }
                    result = result + "</ul>\r\n";
                }
            }
            return result;
        }

        public string GetTree(string MenuName, string OpenMenuName, string StyleSheetPrefix = "") {
            string result = "";
            string Link;
            int EntryPointer;
            string UcaseMenuName;
            string UsedEntries;
            string Caption;
            string JSString;
            if (iEntryCount > 0) {
                UcaseMenuName = MenuName.ToUpper();
                if (StyleSheetPrefix == "") {
                    StyleSheetPrefix = "ccTree";
                }
                if (true) {
                    EntryPointer = 0;
                    Link = iEntry[EntryPointer].Link;
                    if (Link == "") {
                        Link = "javascript: ;";
                    }
                    UsedEntries = "";
                    for (EntryPointer = 0; EntryPointer <= iEntryCount - 1; EntryPointer++) {
                        if (iEntry[EntryPointer].Name == UcaseMenuName) {
                            Caption = iEntry[EntryPointer].Caption;
                            if (iEntry[EntryPointer].Link != "") {
                                Caption = $"<a href=\"{kmaEncodeHTML(cp, iEntry[EntryPointer].Link)}\">{Caption}</a>";
                            }
                            UsedEntries = $"{UsedEntries},{EntryPointer}";
                            result = ""
                                + $"\r\n<ul class=mktree id=tree{iTreeCount}>\r\n"
                                + $"\r\n <li id=\"{iEntry[EntryPointer].Name}\"><span class=mkc>{Caption}</span>"
                                + "\r\n <ul>\r\n"
                                + GetMKTreeBranch(UcaseMenuName, UsedEntries, 2)
                                + "\r\n </ul>\r\n"
                                + "\r\n</li></ul>\r\n";
                            break;
                        }
                    }
                    if (UsedEntries == "") {
                        result = ""
                            + $"\r\n<ul class=mktree id=tree{iTreeCount}>"
                            + GetMKTreeBranch(UcaseMenuName, UsedEntries, 1)
                            + "\r\n</ul>\r\n";
                    }
                    result += "<script src=/resourcelibrary/mktree.js></script>";
                    result += "<script type=\"text/javascript\">convertTrees();";
                    if (OpenMenuName != "") {
                        JSString = OpenMenuName.ToUpper();
                        JSString = JSString.Replace("\\", "\\\\");
                        JSString = JSString.Replace("\r\n", "\\n");
                        JSString = JSString.Replace("'", "\\'");
                        result = $"{result}expandToItem('tree{iTreeCount}','{JSString}');";
                    }
                    result = result + "</script>";
                    iTreeCount = iTreeCount + 1;
                }
            }
            return result;
        }

        private string GetMKTreeBranch(string ParentName, string UsedEntries, int Depth) {
            string result = "";
            int EntryPointer;
            string iUsedEntries;
            int SubMenuCount;
            string ChildMenu;
            string Caption;
            iUsedEntries = UsedEntries;
            SubMenuCount = 0;
            for (EntryPointer = 0; EntryPointer <= iEntryCount - 1; EntryPointer++) {
                if (iEntry[EntryPointer].ParentName == ParentName) {
                    if (($"{iUsedEntries},".IndexOf($",{EntryPointer},", StringComparison.Ordinal) + 1) == 0) {
                        Caption = iEntry[EntryPointer].Caption;
                        if (iEntry[EntryPointer].OnClick != "" && iEntry[EntryPointer].Link != "") {
                            Caption = $"<a href=\"{kmaEncodeHTML(cp, iEntry[EntryPointer].Link)}\" onClick=\"{iEntry[EntryPointer].OnClick}\">{Caption}</a>";
                        } else if (iEntry[EntryPointer].OnClick != "") {
                            Caption = $"<a href=\"#\" onClick=\"{iEntry[EntryPointer].OnClick}\">{Caption}</a>";
                        } else if (iEntry[EntryPointer].Link != "") {
                            Caption = $"<a href=\"{kmaEncodeHTML(cp, iEntry[EntryPointer].Link)}\">{Caption}</a>";
                        } else {
                            Caption = Caption;
                        }
                        iUsedEntries = $"{iUsedEntries},{EntryPointer}";
                        ChildMenu = GetMKTreeBranch(iEntry[EntryPointer].Name, iUsedEntries, Depth + 1);
                        if (newmode) {
                            if (ChildMenu == "") {
                                result = result
                                    + $"\r\n<li class=mklb id=\"{iEntry[EntryPointer].Name}\" >"
                                    + "<div class=\"mkd\">"
                                    + "<span class=mkb>&nbsp;</span>"
                                    + "</div>"
                                    + Caption
                                    + "</li>";
                            } else {
                                result = result
                                    + $"\r\n<li class=\"mklc\" id=\"{iEntry[EntryPointer].Name}\" >"
                                    + "<div class=\"mkd\" >"
                                    + "<span class=mkb onclick=\"mkClick(this)\">&nbsp;</span>"
                                    + "</div>"
                                    + Caption
                                    + "\r\n<ul>"
                                    + ChildMenu
                                    + "\r\n</ul>"
                                    + "</li>";
                            }
                        } else {
                            if (ChildMenu != "") {
                                ChildMenu = ""
                                    + "\r\n<ul>"
                                    + ChildMenu
                                    + "\r\n</ul>"
                                    + "";
                            }
                            result = result
                                + $"\r\n<li class=mklc id=\"{iEntry[EntryPointer].Name}\">"
                                + Caption
                                + ChildMenu
                                + "\r\n</li>";
                        }
                        SubMenuCount = SubMenuCount + 1;
                    }
                }
            }
            return result;
        }
    }
}
