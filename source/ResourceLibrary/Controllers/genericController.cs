
using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.ResourceLibrary.Controllers {
    public sealed class genericController {
        private genericController() {
        }
        public static DateTime encodeMinDate(DateTime srcDate) {
            DateTime returnDate = srcDate;
            if (srcDate < new DateTime(1900, 1, 1)) {
                returnDate = DateTime.MinValue;
            }
            return returnDate;
        }
        public static string getShortDateString(DateTime srcDate) {
            string returnString = "";
            DateTime workingDate = encodeMinDate(srcDate);
            if (!isDateEmpty(srcDate)) {
                returnString = workingDate.ToShortDateString();
            }
            return returnString;
        }
        public static bool isDateEmpty(DateTime srcDate) {
            return (srcDate < new DateTime(1900, 1, 1));
        }
        public static string getSortOrderFromInteger(int id) {
            return id.ToString().PadLeft(7, '0');
        }
        public static string getDateForHtmlInput(DateTime source) {
            if (isDateEmpty(source)) {
                return "";
            } else {
                return $"{source.Year}-{source.Month.ToString().PadLeft(2, '0')}-{source.Day.ToString().PadLeft(2, '0')}";
            }
        }
        public static string convertToDosPath(string sourcePath) {
            return sourcePath.Replace("/", "\\");
        }
        public static string convertToUnixPath(string sourcePath) {
            return sourcePath.Replace("\\", "/");
        }
        public static string Main_getAddonOption(CPBaseClass cp, string requestName, string ignore) {
            return cp.Doc.GetText(requestName);
        }
        public static bool getBoolean_Main_getAddonOption(CPBaseClass cp, string requestName, string ignore) {
            return cp.Doc.GetBoolean(requestName);
        }
        public static void Main_testpoint(CPBaseClass cp, string message) {
            cp.Site.TestPoint(message);
        }
        public static string KmaEncodeSQLNumber(CPBaseClass cp, int src) {
            return cp.Db.EncodeSQLNumber(src);
        }
        public static string KmaEncodeSQLNumber(CPBaseClass cp, double src) {
            return cp.Db.EncodeSQLNumber(src);
        }
        public static string KmaEncodeSQLText(CPBaseClass cp, string src) {
            return cp.Db.EncodeSQLText(src);
        }
        public static string htmlDiv(string innerHtml, string htmlClass = "", string htmlId = "") {
            string result = "<div ";
            result += string.IsNullOrEmpty(htmlClass) ? "" : $" class=\"{htmlClass}\"";
            result += string.IsNullOrEmpty(htmlId) ? "" : $" id=\"{htmlId}\"";
            return $"{result}>{innerHtml}</div>";
        }
        public static string htmlButton(string value, string htmlClass = "", string htmlId = "", string onClick = "") {
            string result = $"<button name=\"name\" value=\"{value}\"";
            result += string.IsNullOrEmpty(htmlClass) ? "" : $" class=\"{htmlClass}\"";
            result += string.IsNullOrEmpty(htmlId) ? "" : $" id=\"{htmlId}\"";
            result += string.IsNullOrEmpty(onClick) ? "" : $" onClick=\"{onClick}\"";
            return $"{result}>";
        }
        public static string htmlHidden(string htmlName, string htmlValue, string htmlClass = "", string htmlId = "") {
            string result = $"<input type=hidden name=\"{htmlName}\" value=\"{htmlValue}\"";
            result += string.IsNullOrEmpty(htmlClass) ? "" : $" class=\"{htmlClass}\"";
            result += string.IsNullOrEmpty(htmlId) ? "" : $" id=\"{htmlId}\"";
            return $"{result}>";
        }
        public static string htmlHidden(string htmlName, int htmlValue, string htmlClass = "", string htmlId = "") {
            return htmlHidden(htmlName, htmlValue.ToString(), htmlClass, htmlId);
        }
        public static string adminUrl(CPBaseClass cp) {
            return cp.Site.GetText("adminurl");
        }
        public static string kmaEncodeURL(CPBaseClass cp, string url) {
            return cp.Utils.EncodeUrl(url);
        }
        public static string Main_GetPanel(string src) {
            return htmlDiv(src);
        }
        public static string kmaEncodeJavascript(CPBaseClass cp, string src) {
            return cp.Utils.EncodeJavascript(src);
        }
        public static string kmaAddSpan(string innerHtml, string htmlClass) {
            return $"<span class=\"{htmlClass}\">{innerHtml}</span>";
        }
        public static string Main_GetFormInputCheckBox(string htmlName, bool isChecked) {
            return $"<input type=checkbox name=\"{htmlName}\" value=1 {(isChecked ? " checked" : "")}>";
        }
        public static string Main_GetPanelInput(string innerHtml) {
            return htmlDiv(innerHtml);
        }
        public static string KmaEncodeMissingText(string src, string ignore = "") {
            return src;
        }
        public static bool KmaEncodeMissingBoolean(bool src, bool ignore = false) {
            return src;
        }
        public static bool KmaEncodeMissingBoolean(bool src, string ignore = "") {
            return src;
        }
        public static string kmaEncodeHTML(CPBaseClass cp, string src) {
            return cp.Utils.EncodeHTML(src);
        }
    }
}
