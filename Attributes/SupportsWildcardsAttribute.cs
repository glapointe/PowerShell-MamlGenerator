using System;

namespace Lapointe.PowerShell.MamlGenerator.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class SupportsWildcardsAttribute : Attribute
    {
    }
}
