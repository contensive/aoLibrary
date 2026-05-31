
using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.ResourceLibrary.Models.View {
    public class DesignBlockViewBaseModel {
        public string styleBackgroundImage { get; set; }
        public string styleheight { get; set; }
        public string contentContainerClass { get; set; }
        public string outerContainerClass { get; set; }
        public static T create<T>(CPBaseClass cp, Models.Db.DesignBlockBaseModel settings) where T : DesignBlockViewBaseModel {
            T result = null;
            try {
                Type instanceType = typeof(T);
                result = (T)Activator.CreateInstance(instanceType);
                result.styleheight = encodeStyleHeight(settings.styleheight);
                result.styleBackgroundImage = encodeStyleBackgroundImage(cp, settings.backgroundImageFilename);
                result.outerContainerClass = settings.themeStyleId.Equals(0) ? string.Empty : $" {cp.Content.GetRecordName("Design Block Themes", settings.themeStyleId)}";
                result.contentContainerClass = ""
                    + (settings.asFullBleed ? " container" : string.Empty)
                    + (settings.padTop ? " pt-5" : " pt-0")
                    + (settings.padRight ? " pr-4" : " pr-0")
                    + (settings.padBottom ? " pb-5" : " pb-0")
                    + (settings.padLeft ? " pl-4" : " pl-0");
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
        public static string encodeStyleHeight(string styleheight) {
            if (string.IsNullOrWhiteSpace(styleheight)) {
                return string.Empty;
            }
            return $"overflow:hidden;height:{styleheight}{(double.TryParse(styleheight, out _) ? "px" : string.Empty)};";
        }
        public static string encodeStyleBackgroundImage(CPBaseClass cp, string backgroundImage) {
            if (string.IsNullOrWhiteSpace(backgroundImage)) {
                return string.Empty;
            }
            return $"background-image: url('{cp.Site.FilePath}{backgroundImage}');";
        }
    }
}
