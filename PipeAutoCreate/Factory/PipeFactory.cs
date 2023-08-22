using Autodesk.Revit.DB;
using PipeAutoCreate.Commponent;
using PipeAutoCreate.DataModel;
using PipeAutoCreate.Extension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.Factory
{
    public static class PipeFactory
    {
        public static void CreatePipes(Document doc, List<Piping> pipings)
        {
            foreach (Piping piping in pipings)
            {
                CreatePipe(doc, piping);
            }
        }

        private static void CreatePipe(Document doc, Piping piping)
        {
            var familySymbol = GetPipeFamilySymbol(doc, piping.Specification);

            var pipeCurve = Line.CreateBound(piping.StartPoint.Converter(), piping.EndPoint.Converter());

            var level = doc.OfClass<Level>().FirstOrDefault(x => x.Elevation == 0);

            var pipe = CreateRPipe(doc, pipeCurve, familySymbol, level);
        }

        private static FamilyInstance CreateRPipe(Document doc, Line pipeCurve, FamilySymbol familySymbol, Level level)
        {
            FamilyInstance ins = doc.Create.NewFamilyInstance(pipeCurve, familySymbol, level, Autodesk.Revit.DB.Structure.StructuralType.Beam);
            return ins;
        }

        private static FamilySymbol GetPipeFamilySymbol(Document doc, string name)
        {
            FamilySymbol symbol;

            var circularSymbols = doc.FindSymbols("圆管");
            var qudrateSymbols = doc.FindSymbols("方管");

            var spc = TextProcessor.ExtractNumbers(name);

            if (spc.Count == 1)
            {
                symbol = circularSymbols.FirstOrDefault(x => x.Name == name);

                if (symbol == null)
                {
                    symbol = circularSymbols.FirstOrDefault().Duplicate(name) as FamilySymbol;
                    symbol.SetParameter("b", spc[0] / 304.8);
                }
            }
            else if (spc.Count == 2)
            {
                symbol = qudrateSymbols.FirstOrDefault(x => x.Name == name);
                if (symbol == null)
                {
                    symbol = qudrateSymbols.FirstOrDefault().Duplicate(name) as FamilySymbol;
                    symbol.SetParameter("b", spc[0] / 304.8);
                    symbol.SetParameter("h", spc[1] / 304.8);
                }
            }
            else
            {
                return null;
            }

            return symbol;
        }
    }
}
