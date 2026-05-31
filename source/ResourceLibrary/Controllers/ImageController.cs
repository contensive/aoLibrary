
using System;
using System.Drawing;
using System.Drawing.Imaging;
using Contensive.BaseClasses;

namespace Contensive.Addons.ResourceLibrary.Controllers {
    public class ImageEditController : IDisposable {
        private bool loaded = false;
        private string src = "";
        private System.Drawing.Image srcImage;
        private int setWidth = 0;
        private int setHeight = 0;
        protected bool disposed = false;
        protected virtual void Dispose(bool disposing) {
            if (!this.disposed) {
                if (disposing) {
                    if (loaded) {
                        srcImage.Dispose();
                        srcImage = null;
                    }
                }
            }
            this.disposed = true;
        }
        public bool load(CPBaseClass cp, string pathFilename) {
            bool returnOk = false;
            try {
                if (cp.File.fileExists(pathFilename)) {
                    src = pathFilename;
                    srcImage = System.Drawing.Image.FromFile($"{cp.Site.PhysicalFilePath}{pathFilename}");
                    setWidth = srcImage.Width;
                    setHeight = srcImage.Height;
                    loaded = true;
                }
            } catch (Exception) {
            }
            return returnOk;
        }
        public bool save(CPBaseClass cp, string pathFilename) {
            bool returnOk = false;
            try {
                if (loaded) {
                    if (src == pathFilename) {
                        if (cp.File.fileExists($"{cp.Site.PhysicalFilePath}{pathFilename}")) {
                            cp.File.DeleteVirtual(pathFilename);
                        }
                    }
                    using (Bitmap imgOutput = new Bitmap(srcImage, setWidth, setHeight)) {
                        ImageFormat imgFormat = srcImage.RawFormat;
                        imgOutput.Save($"{cp.Site.PhysicalFilePath}{pathFilename}", imgFormat);
                    }
                    returnOk = true;
                }
            } catch (Exception) {
            }
            return returnOk;
        }
        public int width {
            get {
                return setWidth;
            }
            set {
                setWidth = value;
            }
        }
        public int height {
            get {
                return setWidth;
            }
            set {
                setHeight = value;
            }
        }
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        ~ImageEditController() {
            Dispose(false);
        }
    }
}
