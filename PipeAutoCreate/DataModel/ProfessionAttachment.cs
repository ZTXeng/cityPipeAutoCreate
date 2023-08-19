using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.DataModel
{
    public class ProfessionAttachment
    {
        public string Profession {  get; set; }

        public List<BuildingAttachment> BuildingAttachments { get; set; }

        public ProfessionAttachment()
        {
            BuildingAttachments = new List<BuildingAttachment>();
        }
    }
}
