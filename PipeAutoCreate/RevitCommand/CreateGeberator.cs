using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PipeAutoCreate.Environment;
using PipeAutoCreate.View;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PipeAutoCreate.RevitCommand
{
    [Transaction(TransactionMode.Manual)]
    public class CreateGeberator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;

            //RevitInfo.SetDocument(doc);

            var view = new ExcelShowView(doc);

            var mainUI = new System.Windows.Interop.WindowInteropHelper(view);
            mainUI.Owner = Process.GetCurrentProcess().MainWindowHandle;
            view.ShowDialog();

            return Result.Succeeded;
        }
    }
}
