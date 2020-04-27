


public struct MenuEntryType {
    public string Caption;           // What is displayed for this entry (does not need to be unique)
    public string Name;              // Unique name for this entry
    public string ParentName;        // Unique name of the parent entry
    public string Link;              // URL
    public string Image;             // Image
    public string ImageOver;         // Image Over
    public string ImageOpen;         // Image when menu is open
    public bool NewWindow;        // True opens link in a new window
    public string OnClick;           // Holds action for onClick
}
