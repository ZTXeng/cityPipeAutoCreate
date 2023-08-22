using MediatR;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PipeAutoCreate.DataModel;
using PipeAutoCreate.ExcelControl;
using PipeAutoCreate.Extension;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DateTest
{
    internal class Program
    {
        private readonly IMediator _mediator;

        private static  IServiceProvider _serviceProvider;

        public Program(IMediator mediator)
        {
            _mediator = mediator;
        }

        static async Task Main(string[] args)
        {
            IServiceCollection cc = new ServiceCollection();
            _serviceProvider = new MyServiceProvider(cc);
            var mdierR = new Mediator(_serviceProvider);
            await mdierR.Send(new MyRequest());
        }
    }
}

