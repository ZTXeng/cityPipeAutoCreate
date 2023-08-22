using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitTest
{
    public static class Class2
    {
        public static IEnumerable<Ttype> OfClass<Ttype>(this Document doc)
        {
            var eles = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(Ttype))
                  .Cast<Ttype>();
            return eles;
        }

        public static void SetParameter(this Element ele,string name, double para)
        {
            var par = ele.LookupParameter(name);

            if (par != null)
            {

                par.Set(para);
            }
        }

        public static List<FamilySymbol> FindSymbols(this Document doc,string name)
        {
            var symbols = new List<FamilySymbol>(); 

            var familiy = doc.OfClass<Family>().FirstOrDefault(f => f.Name == name);

            if (familiy != null)
            {
                var ids = familiy.GetFamilySymbolIds();
                foreach (var id in ids) {
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
