using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace GeoAddin
{
    public abstract class WipeOption
    {
        internal abstract int Execute(string args = null);
    }
}
