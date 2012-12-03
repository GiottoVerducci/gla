using System;

namespace GLA
{
    public abstract class ArmyPart
    {
        public Army Army { get; set; }

        [ExcelData(ColumnName = "index")]
        public string __Index { get { return Id.ToString(); } set { this.Id = Convert.ToInt32(value); } }
        
        public int Id { get; set; }

        public abstract void Resolve();
    }
}