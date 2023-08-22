using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using NPOI.OpenXmlFormats.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.Environment
{
    public static class RevitInfo
    {
        private static UIDocument UIDocument { get; set; }

        private static Document Document { get; set; }

        public static void SetDocument(Document document)
        {
            Document = document;
        }

        public static Document GetDocument()
        {
            return Document;
        }
    }
}
