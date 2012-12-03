using System.Diagnostics;

namespace GLA
{
    [DebuggerDisplay("Famille {Id} {_name}")]
    public class Family : ArmyPart
    {
        [ExcelData(ColumnName = "nom")]
        public string _name { get; set; }

        public override void Resolve()
        {
        }
    }
}