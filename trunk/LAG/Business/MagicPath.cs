using System.Diagnostics;

namespace GLA
{
    [DebuggerDisplay("Voie magique {Id} {_name}")]
    public class MagicPath : ArmyPart
    {
        [ExcelData(ColumnName = "nom")]
        public string _name { get; set; }

        [ExcelData(ColumnName = "qualificatif")]
        public string _qualifier { get; set; }

        public string QualifiedName { get { return _qualifier.Replace("~", _name); } }

        public override void Resolve()
        {
        }
    }
}