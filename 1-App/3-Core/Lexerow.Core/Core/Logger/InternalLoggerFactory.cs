using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Logger;
internal class InternalLoggerFactory : ILoggerFactory
{
    public void AddProvider(ILoggerProvider provider)
    {
        throw new NotImplementedException();
    }

    public ILogger CreateLogger(string categoryName)
    {
        throw new NotImplementedException();
    }

    public void Dispose()
    {
        throw new NotImplementedException();
    }
}
