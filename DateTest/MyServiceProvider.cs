using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DateTest
{
    public class MyServiceProvider : IServiceProvider
    {
        private readonly IServiceCollection services;

        public MyServiceProvider(IServiceCollection services)
        {
            this.services = services;
        }

        public object GetService(Type serviceType)
        {
            return null;
        }
    }
}
