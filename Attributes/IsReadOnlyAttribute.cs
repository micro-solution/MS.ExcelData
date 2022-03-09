using System;

namespace MS.ExcelData.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class IsReadOnlyAttribute : Attribute
    {
        public IsReadOnlyAttribute() { }
    }
}
