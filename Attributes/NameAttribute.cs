using System;

namespace MS.ExcelData.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class NameAttribute : Attribute
    {
        public string Name { get; set; }
        public NameAttribute(string name)
        {
            Name = name;
        }
    }
}
