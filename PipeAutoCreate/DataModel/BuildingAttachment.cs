using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.DataModel
{
    public class BuildingAttachment
    {
        public string Id { get; set; }

        public DPoint Point { get; set; }

        public string Profession { set; get; }

        public string TypeName { get; set; }

        //特征
        public string Characteristic { get; set; }

        //规格
        public string Specification { get; set; }

        //地面高程
        public double GroundHegiht { get; set; }

        //埋深
        public double BuridHegiht { get; set; }

        public List<ExtensionParameter> ExtensionParameters { get; set; }

        public BuildingAttachment()
        {
            ExtensionParameters = new List<ExtensionParameter>();
        }
    }
}
