using MediatR;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.Request
{
    public class CreatePipeRequest:IRequest<bool>
    {
        public string PipeName { get; set; }

        public CreatePipeRequest()
        {
                
        }
    }
}
