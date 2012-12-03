using System.Collections.Generic;
using System.Diagnostics;

namespace GLA
{
    [DebuggerDisplay("Capacity {Id} {_name}")]
    public class Capacity : ArmyPart
    {
        [ExcelData(ColumnName = "nom")]
        public string _name { get; set; }

        [ExcelData(ColumnName = "description")]
        public string _description { get; set; }

        public readonly List<Unit> AssociatedUnits = new List<Unit>();

        private Dictionary<string, List<Unit>> _groupedAssociatedUnits;
        public Dictionary<string, List<Unit>> GroupedAssociatedUnits
        {
            get
            {
                return _groupedAssociatedUnits ?? (_groupedAssociatedUnits = Unit.GetGroupedUnits(this.AssociatedUnits));
            }
        }

        public override void Resolve()
        {
            
        }
    }
}