
using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.ResourceLibrary.Controllers {
    public class InstanceIdController {
        public static string getSettingsGuid(CPBaseClass cp, string designBlockName, ref string returnHtmlMessage) {
            if (string.IsNullOrWhiteSpace(designBlockName)) {
                throw new ApplicationException("getInstanceId called without valid designBlockName.");
            }
            string result = cp.Doc.GetText("instanceId");
            if (!string.IsNullOrWhiteSpace(result)) {
                return result;
            }
            result = cp.Doc.GetText("forminstanceId");
            if (!string.IsNullOrWhiteSpace(result)) {
                return result;
            }
            if (cp.Doc.PageId > 0) {
                result = $"DesignBlockUsedWithoutInstanceId-[{designBlockName}]-PageId-{cp.Doc.PageId}";
                if (!string.IsNullOrEmpty(cp.Doc.GetText(result))) {
                    returnHtmlMessage += "<p>Error, this design block is used twice on this page. This is only allowed if it was added with the drag-drop tool, or includes a unique instance id.</p>";
                    cp.Site.ErrorReport($"Design Block [{designBlockName}] on page [#{cp.Doc.PageId},{cp.Doc.PageName}] does not include an instanceId and was used on the page twice.");
                    return string.Empty;
                }
                cp.Doc.SetProperty(result, "used");
                return result;
            }
            if (cp.Request.PathPage == cp.Site.GetText("adminurl")) {
                result = $"DesignBlockUsedOnAdminSite-[{designBlockName}]";
                if (!string.IsNullOrEmpty(cp.Doc.GetText(result))) {
                    returnHtmlMessage += "<p>Error, this design block is used twice on the admin site.</p>";
                    cp.Site.ErrorReport($"Design Block [{designBlockName}] on page [#{cp.Doc.PageId},{cp.Doc.PageName}] does not include an instanceId and was used on the page twice.");
                    return string.Empty;
                }
                return result;
            }
            throw new ApplicationException($"Design Block [{designBlockName}] called without instanceId must be on a page or the admin site.");
        }
    }
}
