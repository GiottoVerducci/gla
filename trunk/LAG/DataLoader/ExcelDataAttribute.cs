using System;

namespace GLA
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ExcelDataAttribute : Attribute
    {
        public string ColumnName { get; set; }
    }
}