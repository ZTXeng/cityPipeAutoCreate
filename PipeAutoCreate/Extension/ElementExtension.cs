using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.Extension
{
    public static class ElementExtension
    {
        public static void SetParameter(this Element ele, string name, double para)
        {
            var par = ele.LookupParameter(name);

            if (par != null)
            {

                par.Set(para);
            }
        }
    }
}
