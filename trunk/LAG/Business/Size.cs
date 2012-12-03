using System.Diagnostics;

namespace GLA
{
    [DebuggerDisplay("Taille {Id} {_name}")]
    public class Size : ArmyPart
    {
        [ExcelData(ColumnName = "nom")]
        public string _name { get; set; }

        public override void Resolve()
        {
        }
    }
}