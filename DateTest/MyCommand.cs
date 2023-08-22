using MediatR;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DateTest
{
    internal class MyCommand : IRequestHandler<MyRequest, bool>
    {
        Task<bool> IRequestHandler<MyRequest, bool>.Handle(MyRequest request, CancellationToken cancellationToken)
        {


           return Task.FromResult(true);
        }
    }
}
