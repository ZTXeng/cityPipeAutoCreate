using MediatR;
using PipeAutoCreate.Request;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PipeAutoCreate.Command
{
    public class CreatePipeCommand : IRequestHandler<CreatePipeRequest, bool>
    {
        Task<bool> IRequestHandler<CreatePipeRequest, bool>.Handle(CreatePipeRequest request, CancellationToken cancellationToken)
        {
            Console.WriteLine("Hellow");


            return Task.FromResult(true);
        }
    }
}
