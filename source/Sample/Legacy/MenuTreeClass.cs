
using System;
using System.Collections.Generic;
using Contensive.BaseClasses;

namespace Contensive.Addons.ResourceLibrary {
    // 
    public class menuTreeClass {
        // 
        private CPBaseClass cp;

        internal menuTreeClass(CPBaseClass cp) {
            this.cp = cp;
            EntryIndexName = new FastIndex5Class();
            VBMath.Randomize();
            MenuFlyoutNamePrefix = "id" + System.Convert.ToString(Conversion.Int(9999 * VBMath.Rnd()));
        }
        // 
        // ----- Each menu item has an MenuEntry
        // 

        public string iMenuFilePath;
        // 
        // ----- Menu Entry storage
        // 
        public int iEntryCount;           // Count of Menus in the object
        public int iEntrySize;
        public MenuEntryType[] iEntry;
        // 
        // Private iDQMCount as integer            ' Count of Default Menus for this instance
        // Private iDQMCLosed As Boolean       ' true if the menu has been closed
        // 
        public int iTreeCount;           // Count of Tree Menus for this instance
        public string iMenuCloseString;  // String returned for closing menus
        // 
        public List<string> UsedEntries = new List<string>();
        public FastIndexClass EntryIndexName;
        public FastIndexClass EntryIndexID;
        // 
        // ----- RollOverFlyout storage
        // 
        public string MenuFlyoutNamePrefix;    // Random prefix added to element IDs to avoid namespace collision
        public string MenuFlyoutIcon_Local;      // string used to mark a button that has a non-hover flyout
        public const bool newmode = true;
        // 
        // ===============================================================================
        // Create a new Menu Entry
        // ===============================================================================
        // 
        public void AddEntry(string EntryName, string ParentiEntryName, string ImageLink, string ImageOverLink, string Link, string Caption, string OnClickJavascript = "", string Ignore1 = "", string ImageOpenLink = "", bool NewWindow = false) {
            string iEntryName;
            string UcaseEntryName;
            // 
            iEntryName = EntryName.Replace(",", "");
            UcaseEntryName = iEntryName.ToUpperInvariant();
            // 
            if (string.IsNullOrWhiteSpace(iEntryName)) { return; }
            if (UsedEntries.Contains(UcaseEntryName)) { return; }
            UsedEntries.Add(UcaseEntryName);

            if ((iEntryName != "") & (Strings.InStr(1, UsedEntries + ",", "," + UcaseEntryName + ",", Constants.vbBinaryCompare) == 0)) {
                UsedEntries = UsedEntries + "," + UcaseEntryName;
                if (iEntryCount >= iEntrySize) {
                    iEntrySize = iEntrySize + 10;
                    var oldIEntry = iEntry;
                    iEntry = new MenuEntryType[iEntrySize + 1];
                    if (oldIEntry != null)
                        Array.Copy(oldIEntry, iEntry, Math.Min(iEntrySize + 1, oldIEntry.Length));
                }
                MenuEntryType withBlock = iEntry[iEntryCount];
                withBlock.Link = Link;
                withBlock.Image = ImageLink;
                withBlock.OnClick = OnClickJavascript;
                withBlock.Name = UcaseEntryName;
                withBlock.ParentName = ParentiEntryName.ToUpperInvariant();
                withBlock.ImageOver = ImageOverLink;
                withBlock.ImageOpen = ImageOpenLink;
                withBlock.NewWindow = NewWindow;
                withBlock.Caption = string.IsNullOrWhiteSpace(Caption) ? iEntryName : Caption;
                EntryIndexName.SetPointer(UcaseEntryName, iEntryCount);
                iEntryCount = iEntryCount + 1;
            }
        }
        // 
        // ===============================================================================
        // Returns the menu specified, if it is in local storage
        // 
        // It also creates the menu data in a close string that is returned in GetMenuClose.
        // It must be done there so the link buttons height can be calculated.
        // ===============================================================================
        // 
        public string GetMenu(string MenuName, string StyleSheetPrefix = "") {
            GetMenu = GetTree(MenuName, "", KmaEncodeMissingText(StyleSheetPrefix, "ccTree"));
        }
        // 
        // ===============================================================================
        // Gets the Menu Branch for the Tree Menu
        // ===============================================================================
        // 
        private string GetMenuTreeBranch(string ParentName, string JSObject, string UsedEntries) {
            // 
            int EntryPointer;
            string iUsedEntries;
            string JSChildObject;
            int SubMenuCount;
            // 
            iUsedEntries = UsedEntries;
            SubMenuCount = 0;
            for (EntryPointer = 0; EntryPointer <= iEntryCount - 1; EntryPointer++) {
                {
                    var withBlock = iEntry[EntryPointer];
                    if (withBlock.ParentName == ParentName) {
                        if ((Strings.InStr(1, iUsedEntries + ",", "," + EntryPointer + ",") == 0)) {
                            JSChildObject = JSObject + ".s[" + SubMenuCount + "]";
                            iUsedEntries = iUsedEntries + "," + EntryPointer;
                            GetMenuTreeBranch = GetMenuTreeBranch
                        + JSChildObject + " = new so(0,'" + withBlock.Caption + "','" + withBlock.Link + "','_blank',''); " + Constants.vbCrLf
                        + GetMenuTreeBranch(withBlock.Name, JSChildObject, iUsedEntries);
                            SubMenuCount = SubMenuCount + 1;
                        }
                    }
                }
            }
        }
        // 
        // ===============================================================================
        // Returns the menu specified, if it is in local storage
        // 
        // It also creates the menu data in a close string that is returned in GetMenuClose.
        // It must be done there so the link buttons height can be calculated.
        // Uses a simple UL/Stylesheet method, returning to the server with every click
        // ===============================================================================
        // 
        private string GetMenuTreeList(string MenuName, string OpenNodesList) {
            // 
            int EntryPointer;
            string UcaseMenuName;
            // 
            // ----- Get the menu pointer
            // 
            if (iEntryCount > 0) {
                UcaseMenuName = Strings.UCase(MenuName);
                EntryPointer = EntryIndexName.GetPointer(UcaseMenuName);
                GetMenuTreeList = GetMenuTreeListBranch2(EntryPointer, "", OpenNodesList);
                return;
            }
        }
        // 
        // ===============================================================================
        // Gets the Menu Branch for the Tree Menu
        // ===============================================================================
        // 
        private string GetMenuTreeListBranch2(int NodePointer, string UsedEntriesList, string OpenNodesList) {
            // 
            string Link;
            int EntryPointer;
            string UcaseNodeName;
            string Image;
            string Caption;
            // 
            if (iEntryCount > 0) {
                // 
                // Output this node
                // 
                if (Strings.InStr(1, "," + System.Convert.ToString(NodePointer) + ",", "," + UsedEntriesList + ",") == 0) {
                    GetMenuTreeListBranch2 = GetMenuTreeListBranch2 + "<ul Style=\"list-style-type: none; margin-left: 20px\">";
                    // 
                    // The Node has not already been used in this branch
                    // 
                    Caption = iEntry[NodePointer].Caption;
                    Link = kmaEncodeHTML(cp, iEntry[NodePointer].Link);
                    if (Link != "")
                        Caption = "<A TARGET=\"_blank\" HREF=\"" + Link + "\">" + Caption + "</A>";
                    // 
                    if (Strings.InStr(1, "," + OpenNodesList + ",", "," + System.Convert.ToString(NodePointer) + ",") == 0) {
                        // 
                        // The branch is closed
                        // 
                        Image = iEntry[NodePointer].Image;
                        GetMenuTreeListBranch2 = GetMenuTreeListBranch2 + "<li><A HREF=\"?OpenNodesList=" + OpenNodesList + "&OpenNode=" + NodePointer + "\"><IMG SRC=\"" + Image + "\" HEIGHT=\"18\" WIDTH=\"18\" BORDER=0 ALT=\"Open Folder\" /></A>&nbsp;" + Caption + "</li>";
                    } else {
                        // 
                        // The branch is open
                        // 
                        Image = iEntry[NodePointer].ImageOpen;
                        if (Image == "")
                            Image = iEntry[NodePointer].Image;
                        GetMenuTreeListBranch2 = GetMenuTreeListBranch2
                    + "<li>"
                    + "<A HREF=\"?OpenNodesList=" + OpenNodesList + "&CloseNode=" + NodePointer + "\">"
                    + "<IMG SRC=\"" + Image + "\" HEIGHT=\"18\" WIDTH=\"18\" BORDER=0 ALT=\"Close Folder\" />"
                    + "</A>&nbsp;" + Caption + "</li>";
                        // 
                        // Now output any child branches of this node
                        // 
                        UcaseNodeName = Strings.UCase(iEntry[NodePointer].Name);
                        for (EntryPointer = 0; EntryPointer <= iEntryCount - 1; EntryPointer++) {
                            if ((iEntry[EntryPointer].ParentName == UcaseNodeName))
                                GetMenuTreeListBranch2 = GetMenuTreeListBranch2 + GetMenuTreeListBranch2(EntryPointer, UsedEntriesList + "," + NodePointer, OpenNodesList);
                        }
                    }
                    GetMenuTreeListBranch2 = GetMenuTreeListBranch2 + "</ul>" + Constants.vbCrLf;
                }
            }
        }
        // 
        // ===============================================================================
        // Returns the menu specified, if it is in local storage
        // 
        // It also creates the menu data in a close string that is returned in GetTreeClose.
        // It must be done there so the link buttons height can be calculated.
        // ===============================================================================
        // 
        public string GetTree(string MenuName, string OpenMenuName, string StyleSheetPrefix = "") {
            GetTree = "";
            string Link;
            int EntryPointer;
            string UcaseMenuName;
            string UsedEntries;
            string Caption;
            string JSString;
            // 
            // ----- Get the menu pointer
            // 
            if (iEntryCount > 0) {
                UcaseMenuName = Strings.UCase(MenuName);
                if (StyleSheetPrefix == "")
                    StyleSheetPrefix = "ccTree";
                if (true) {
                    // 
                    // ----- Build the linked -button-
                    // 
                    Link = iEntry[EntryPointer].Link;
                    if (Link == "")
                        Link = "javascript: ;";
                    // 
                    // Find the Menu Entry, and create the top element here
                    // 
                    UsedEntries = "";
                    for (EntryPointer = 0; EntryPointer <= iEntryCount - 1; EntryPointer++) {
                        {
                            var withBlock = iEntry[EntryPointer];
                            if (withBlock.Name == UcaseMenuName) {
                                Caption = withBlock.Caption;
                                if (withBlock.Link != "")
                                    Caption = "<a href=\"" + kmaEncodeHTML(cp, withBlock.Link) + "\">" + Caption + "</a>";
                                UsedEntries = UsedEntries + "," + System.Convert.ToString(EntryPointer);
                                GetTree = ""
                            + Constants.vbCrLf + "<ul class=mktree id=tree" + iTreeCount + ">" + Constants.vbCrLf
                            + Constants.vbCrLf + " <li id=\"" + withBlock.Name + "\"><span class=mkc>" + Caption + "</span>"
                            + Constants.vbCrLf + " <ul>" + Constants.vbCrLf
                            + GetMKTreeBranch(UcaseMenuName, UsedEntries, 2)
                            + Constants.vbCrLf + " </ul>" + Constants.vbCrLf
                            + Constants.vbCrLf + "</li></ul>" + Constants.vbCrLf;
                                break;
                            }
                        }
                    }
                    if (UsedEntries == "")
                        GetTree = ""
+ Constants.vbCrLf + "<ul class=mktree id=tree" + iTreeCount + ">"
+ GetMKTreeBranch(UcaseMenuName, UsedEntries, 1)
+ Constants.vbCrLf + "</ul>" + Constants.vbCrLf;
                    GetTree += "<script src=/resourcelibrary/mktree.js></script>";
                    GetTree += "<script type=\"text/javascript\">convertTrees();";
                    if (OpenMenuName != "") {
                        JSString = Strings.UCase(OpenMenuName);
                        JSString = Strings.Replace(JSString, @"\", @"\\");
                        JSString = Strings.Replace(JSString, Constants.vbCrLf, @"\n");
                        JSString = Strings.Replace(JSString, "'", @"\'");
                        // Call Main.AddOnLoadJavascript("expandToItem('tree" & iTreeCount & "','" & JSString & "');")
                        GetTree = GetTree + "expandToItem('tree" + iTreeCount + "','" + JSString + "');";
                    }
                    GetTree = GetTree + "</script>";
                    // 
                    // increment the menu count
                    // 
                    iTreeCount = iTreeCount + 1;
                }
            }
        }
        // 
        // ===============================================================================
        // Gets the Menu Branch for the Tree Menu
        // ===============================================================================
        // 
        private string GetMKTreeBranch(string ParentName, string UsedEntries, int Depth) {
            // 
            int EntryPointer;
            string iUsedEntries;
            int SubMenuCount;
            string ChildMenu;
            string Caption;
            // 
            iUsedEntries = UsedEntries;
            SubMenuCount = 0;
            for (EntryPointer = 0; EntryPointer <= iEntryCount - 1; EntryPointer++) {
                {
                    var withBlock = iEntry[EntryPointer];
                    if (withBlock.ParentName == ParentName) {
                        if ((Strings.InStr(1, iUsedEntries + ",", "," + EntryPointer + ",") == 0)) {
                            Caption = withBlock.Caption;
                            if (withBlock.OnClick != "" & withBlock.Link != "")
                                Caption = "<a href=\"" + kmaEncodeHTML(cp, withBlock.Link) + "\" onClick=\"" + withBlock.OnClick + "\">" + Caption + "</a>";
                            else if (withBlock.OnClick != "")
                                Caption = "<a href=\"#\" onClick=\"" + withBlock.OnClick + "\">" + Caption + "</a>";
                            else if (withBlock.Link != "")
                                Caption = "<a href=\"" + kmaEncodeHTML(cp, withBlock.Link) + "\">" + Caption + "</a>";
                            else
                                Caption = Caption;
                            iUsedEntries = iUsedEntries + "," + EntryPointer;

                            ChildMenu = GetMKTreeBranch(withBlock.Name, iUsedEntries, Depth + 1);
                            if (newmode) {
                                if (ChildMenu == "")
                                    GetMKTreeBranch = GetMKTreeBranch
+ Constants.vbCrLf + "<li class=mklb id=\"" + withBlock.Name + "\" >"
+ "<div class=\"mkd\">"
+ "<span class=mkb>&nbsp;</span>"
+ "</div>"
+ Caption
+ "</li>";
                                else
                                    // 
                                    // 3/18/2010 changes to keep firefox from blocking clicks
                                    // 
                                    GetMKTreeBranch = GetMKTreeBranch
                                + Constants.vbCrLf + "<li class=\"mklc\" id=\"" + withBlock.Name + "\" >"
                                + "<div class=\"mkd\" >"
                                + "<span class=mkb onclick=\"mkClick(this)\">&nbsp;</span>"
                                + "</div>"
                                + Caption
                                + Constants.vbCrLf + "<ul>"
                                + ChildMenu
                                + Constants.vbCrLf + "</ul>"
                                + "</li>";
                            } else {
                                if (ChildMenu != "")
                                    ChildMenu = ""
+ Constants.vbCrLf + "<ul>"
+ ChildMenu
+ Constants.vbCrLf + "</ul>"
+ "";
                                GetMKTreeBranch = GetMKTreeBranch
                            + Constants.vbCrLf + "<li class=mklc id=\"" + withBlock.Name + "\">"
                            + Caption
                            + ChildMenu
                            + Constants.vbCrLf + "</li>";
                            }
                            SubMenuCount = SubMenuCount + 1;
                        }
                    }
                }
            }
        }
    }
}
