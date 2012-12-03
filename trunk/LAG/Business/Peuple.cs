using System.Collections.Generic;
using System.Diagnostics;

namespace GLA
{
    [DebuggerDisplay("Peuple {Id} {_name}")]
    public class Peuple : ArmyPart
    {
        [ExcelData(ColumnName = "nom")]
        public string _name { get; set; }

        [ExcelData(ColumnName = "qualificatif")]
        public string _qualifier { get; set; }

        public readonly List<MagicPath> AssociatedMagicPaths = new List<MagicPath>();

        public override void Resolve()
        {
        }
    }
}