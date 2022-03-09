using System;

namespace MS.ExcelData.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class IsIndexAttribute : Attribute
    {
        public IsIndexAttribute() { }
    }
}
