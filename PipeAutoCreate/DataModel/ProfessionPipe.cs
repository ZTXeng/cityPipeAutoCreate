using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.DataModel
{
    public class ProfessionPipe
    {
        public string PProfession { get; set; }

        public List<Piping> Pipings { get; set; }

        public ProfessionPipe()
        {
            Pipings = new List<Piping>();
        }
    }
}
