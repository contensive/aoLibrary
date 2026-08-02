
using Contensive.Addons.ResourceLibrary.Controllers;
using Contensive.Addons.ResourceLibrary.Models.View;
using Contensive.BaseClasses;
using static Contensive.Addons.ResourceLibrary.constants;

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
                if (CP.Doc.IsAdminSite) {
                    //
                    // -- admin site, wrap in LayoutBuilder for portal styling
                    var legacyLibrary = new LegacyLibraryClass();
                    var htmlBody = legacyLibrary.getResourceLibrary(CP, useLayoutBuilder: true);
                    //
                    var form = CP.AdminUI.CreateLayoutBuilder();
                    form.title = designBlockName;
                    form.description = "Manage files and folders used throughout the site.";
                    form.isOuterContainer = true;
                    form.body = htmlBody;
                    //
                    // -- add buttons matching the legacy button bar
                    form.addFormButton(ButtonCancel);
                    form.addFormButton(ButtonApply);
                    if (legacyLibrary.hasModifyAccess) {
                        form.addFormButton(ButtonDelete);
                    }
                    return form.getHtml();
                }
                //
                // -- public site design block settings
                string userErrorMessage = "";
                var settingsGuid = InstanceIdController.getSettingsGuid(CP, designBlockName, ref userErrorMessage);
                if (string.IsNullOrEmpty(settingsGuid)) {
                    return userErrorMessage;
                }
                //
                // -- locate or create a data record for this guid
                var publicSettings = Models.Db.ResourceLibraryModel.createOrAddSettings(CP, settingsGuid);
                if (publicSettings == null) {
                    throw new System.ApplicationException("Could not create the design block settings record.");
                }
                //
                CP.Doc.SetProperty("RootFolderName", publicSettings.RootFolderName);
                CP.Doc.SetProperty("Block Folder Navigation", publicSettings.BlockFolderNavigation);
                var publicHtmlBody = (new LegacyLibraryClass()).getResourceLibrary(CP);
                //
                // -- translate the Db model to a view model and mustache it into the layout
                var viewModel = LibraryViewModel.create(CP, publicSettings, publicHtmlBody);
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
