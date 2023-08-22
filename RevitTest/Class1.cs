using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace RevitTest
{
    [Transaction(TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;

            var tran = new Transaction(doc);

            tran.Start("d");

            var spc = "55";

            FamilySymbol familySymbol;

            var pipeCurve = Line.CreateBound(XYZ.Zero, new XYZ(1000, 0, 0));

            var circularSymbols = doc.FindSymbols("圆管");
            var qudrateSymbols = doc.FindSymbols("方管");

            familySymbol = circularSymbols.FirstOrDefault(x => x.Name.Equals("55"));

            if (familySymbol == null)
            {
                familySymbol = circularSymbols.FirstOrDefault().Duplicate("55") as FamilySymbol;
                familySymbol.SetParameter("b",55/304.8);
            }

            var level = doc.OfClass<Level>().First();

            FamilyInstance ins = doc.Create.NewFamilyInstance(pipeCurve, familySymbol, level, Autodesk.Revit.DB.Structure.StructuralType.Beam);

            tran.Commit();

            return Result.Succeeded;
        }

    }
}
