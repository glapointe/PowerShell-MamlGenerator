using System;

namespace Lapointe.PowerShell.MamlGenerator.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public class ExampleAttribute : Attribute
    {
        public string Code { get; set; }
        public string Remarks { get; set; }
    }
}
