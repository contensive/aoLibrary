
using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.ResourceLibrary.Models.View {
    public class LibraryViewModel : DesignBlockViewBaseModel {
        public string bodyHtml { get; set; }
        public static LibraryViewModel create(CPBaseClass cp, Models.Db.ResourceLibraryModel settings, string htmlBody) {
            try {
                var result = DesignBlockViewBaseModel.create<LibraryViewModel>(cp, settings);
                result.bodyHtml = htmlBody;
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return null;
            }
        }
    }
}
