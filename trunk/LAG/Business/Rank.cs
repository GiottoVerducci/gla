using System.Diagnostics;

namespace GLA
{
    [DebuggerDisplay("Rank {Id} {_name}")]
    public class Rank : ArmyPart
    {
        [ExcelData(ColumnName = "nom")]
        public string _name { get; set; }

        public override void Resolve()
        {
        }
    }
}