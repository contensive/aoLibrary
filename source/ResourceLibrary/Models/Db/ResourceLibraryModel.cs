
using Contensive.BaseClasses;
using Contensive.Models.Db;

namespace Contensive.Addons.ResourceLibrary.Models.Db {
    public class ResourceLibraryModel : DesignBlockBaseModel {
        public static DbBaseTableMetadataModel tableMetadata { get; } = new DbBaseTableMetadataModel("Resource Libraries", "ccResourceLibraries", "default", false);
        public string RootFolderName { get; set; }
        public bool BlockFolderNavigation { get; set; }
        public static ResourceLibraryModel createOrAddSettings(CPBaseClass cp, string settingsGuid) {
            ResourceLibraryModel result = create<ResourceLibraryModel>(cp, settingsGuid);
            if (result == null) {
                result = DesignBlockBaseModel.addDefault<ResourceLibraryModel>(cp);
                result.name = $"{ResourceLibraryModel.tableMetadata.contentName} {result.id}";
                result.ccguid = settingsGuid;
                result.themeStyleId = 0;
                result.padTop = false;
                result.padBottom = false;
                result.padRight = false;
                result.padLeft = false;
                result.RootFolderName = "";
                result.BlockFolderNavigation = false;
                result.save(cp);
                cp.Content.LatestContentModifiedDate.Track(result.modifiedDate);
            }
            return result;
        }
    }
}
