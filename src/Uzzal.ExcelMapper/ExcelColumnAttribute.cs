using System;

namespace Uzzal.ExcelMapper
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute: Attribute
    {
        
        public string Name { get; set; }

        public ExcelColumnAttribute(string name)
        {
            Name = name;
        }
    }
}
