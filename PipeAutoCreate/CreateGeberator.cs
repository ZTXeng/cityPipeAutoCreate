using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PipeAutoCreate
{
    [Transaction(TransactionMode.Manual)]
    public class CreateGeberator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;

            var path = "";
            var dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                path = dialog.FileName;
            }

            if (path == "")  //dialog窗口直接取消时，取消操作
            {
                return Result.Cancelled;
            }

            return Result.Succeeded;
        }
    }
}
