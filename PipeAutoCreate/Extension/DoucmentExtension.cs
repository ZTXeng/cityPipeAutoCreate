using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.Extension
{
    public static class DoucmentExtension
    {
        public static IEnumerable<Ttype> OfClass<Ttype>(this Document doc)
        {
            var eles = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(Ttype))
                  .Cast<Ttype>();
            return eles;
        }

        public static List<FamilySymbol> FindSymbols(this Document doc, string name)
        {
            var symbols = new List<FamilySymbol>();

            var familiy = doc.OfClass<Family>().FirstOrDefault(f => f.Name == name);

            if (familiy != null)
            {
                var ids = familiy.GetFamilySymbolIds();
                foreach (var id in ids)
                {
                    var symbol = doc.GetElement(id) as FamilySymbol;
                    symbols.Add(symbol);
                }
                return symbols;
            }
            else
            {
                return null;
            }
        }
    }
}
