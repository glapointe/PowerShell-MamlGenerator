using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.PowerShell.MamlGenerator.Attributes
{
    public class CmdletGroupAttribute : Attribute
    {
        public string Group { get; set; }
        public CmdletGroupAttribute(string group)
        {
            Group = group;
        }
    }
}
