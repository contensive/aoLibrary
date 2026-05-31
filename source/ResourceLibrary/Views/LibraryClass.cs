
using Contensive.Addons.ResourceLibrary.Controllers;
using Contensive.Addons.ResourceLibrary.Models.View;
using Contensive.BaseClasses;

namespace Contensive.Addons.ResourceLibrary.Views {
    //
    //====================================================================================================
    /// <summary>
    /// Design block with a centered headline, image, paragraph text and a button.
    /// </summary>
    public class LibraryClass : AddonBaseClass {
        //
        //====================================================================================================
        //
        public override object Execute(CPBaseClass CP) {
            const string designBlockName = "Resource Library";
            try {
                Models.Db.ResourceLibraryModel settings;
                if (CP.Doc.IsAdminSite) {
                    //
                    // -- admin site settings
                    settings = new Models.Db.ResourceLibraryModel() {
                        RootFolderName = "",
                        BlockFolderNavigation = false
                    };
                } else {
                    //
                    // -- design block settings
                    string userErrorMessage = "";
                    var settingsGuid = InstanceIdController.getSettingsGuid(CP, designBlockName, ref userErrorMessage);
                    if (string.IsNullOrEmpty(settingsGuid)) {
                        return userErrorMessage;
                    }
                    //
                    // -- locate or create a data record for this guid
                    settings = Models.Db.ResourceLibraryModel.createOrAddSettings(CP, settingsGuid);
                    if (settings == null) {
                        throw new System.ApplicationException("Could not create the design block settings record.");
                    }
                    //
                    CP.Doc.SetProperty("RootFolderName", settings.RootFolderName);
                    CP.Doc.SetProperty("Block Folder Navigation", settings.BlockFolderNavigation);
                }
                var htmlBody = (new LegacyLibraryClass()).getResourceLibrary(CP);
                //
                // -- translate the Db model to a view model and mustache it into the layout
                var viewModel = LibraryViewModel.create(CP, settings, htmlBody);
                if (viewModel == null) {
                    throw new System.ApplicationException("Could not create design block view model.");
                }
                // Read the embedded resource
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using var stream = assembly.GetManifestResourceStream("Contensive.Addons.ResourceLibrary.Resources.LibraryLayout.txt");
                using var reader = new System.IO.StreamReader(stream);
                string layout = reader.ReadToEnd();
                return Nustache.Core.Render.StringToString(layout, viewModel);
            } catch (System.Exception ex) {
                CP.Site.ErrorReport(ex);
                return $"<!-- {designBlockName}, Unexpected Exception -->";
            }
        }
    }
}
