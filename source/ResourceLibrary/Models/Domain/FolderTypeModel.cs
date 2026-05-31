
namespace Contensive.Addons.ResourceLibrary.Models.Domain {
    public class FolderTypeModel {
        public int FolderID;
        public int parentFolderID;
        public string Name;
        public string FullPath;
        public bool hasViewAccess;
        public bool viewAccessIsValid;
        public bool hasModifyAccess;
        public bool modifyAccessIsValid;
    }
}
