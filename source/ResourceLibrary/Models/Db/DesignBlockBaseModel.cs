
using System;
using System.Collections.Generic;
using System.Text;
using Contensive.BaseClasses;

namespace Contensive.Addons.ResourceLibrary.Models.Db {
    public class DesignBlockBaseModel : Contensive.Models.Db.DbBaseModel {
        public string backgroundImageFilename { get; set; }
        public int themeStyleId { get; set; }
        public bool padTop { get; set; }
        public bool padBottom { get; set; }
        public bool padRight { get; set; }
        public bool padLeft { get; set; }
        public string styleheight { get; set; }
        public bool asFullBleed { get; set; }
    }
}
