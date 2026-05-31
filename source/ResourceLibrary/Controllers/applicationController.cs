
using System;
using System.Collections.Generic;
using System.Text.Json;
using Contensive.BaseClasses;

namespace Contensive.Addons.ResourceLibrary.Controllers {
    public class applicationController : IDisposable {
        private CPBaseClass cp;
        public bool allowPlace { get; }
        public string topFolderPath { get; private set; }
        public bool AllowGroupAdd { get; }
        public List<packageErrorClass> packageErrorList { get; set; } = new List<packageErrorClass>();
        public List<packageNodeClass> packageNodeList { get; set; } = new List<packageNodeClass>();
        public List<packageProfileClass> packageProfileList { get; set; } = new List<packageProfileClass>();

        public string getSerializedPackage() {
            string result = "";
            try {
                result = serializeObject(cp, new packageClass {
                    success = packageErrorList.Count.Equals(0),
                    nodeList = packageNodeList,
                    errorList = packageErrorList,
                    profileList = packageProfileList
                });
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }

        public applicationController(CPBaseClass cp) {
            this.cp = cp;
            allowPlace = cp.Doc.GetBoolean("AllowSelectResource");
            topFolderPath = cp.Doc.GetText("RootFolderName");
            topFolderPath = topFolderPath.Trim();
            topFolderPath = topFolderPath.ToLower();
            topFolderPath = topFolderPath.Replace("/", "\\");
            if (topFolderPath.Length >= 4 && topFolderPath.Substring(0, 4) == "root") {
                topFolderPath = topFolderPath.Substring(4);
            }
            if (topFolderPath.Length >= 1 && topFolderPath.Substring(0, 1) == "\\") {
                topFolderPath = topFolderPath.Substring(1);
            }
            if (topFolderPath.Length >= 1 && topFolderPath.Substring(topFolderPath.Length - 1) == "\\") {
                topFolderPath = topFolderPath.Substring(0, topFolderPath.Length - 1);
            }
            AllowGroupAdd = cp.Doc.GetBoolean("AllowGroupAdd");
        }

        public static string serializeObject(CPBaseClass CP, object dataObject) {
            string result = "";
            try {
                result = JsonSerializer.Serialize(dataObject);
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
            }
            return result;
        }

        public class packageProfileClass {
            public string name;
            public int time;
        }

        [Serializable]
        public class packageClass {
            public bool success = false;
            public List<packageErrorClass> errorList = new List<packageErrorClass>();
            public List<packageNodeClass> nodeList = new List<packageNodeClass>();
            public List<packageProfileClass> profileList;
        }

        [Serializable]
        public class packageNodeClass {
            public string dataFor = "";
            public object data;
        }

        [Serializable]
        public class packageErrorClass {
            public int number = 0;
            public string description = "";
        }

        #region  IDisposable Support
        protected bool disposed = false;
        protected virtual void Dispose(bool disposing) {
            if (!this.disposed) {
                if (disposing) {
                }
            }
            this.disposed = true;
        }
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        ~applicationController() {
            Dispose(false);
        }
        #endregion
    }
}
