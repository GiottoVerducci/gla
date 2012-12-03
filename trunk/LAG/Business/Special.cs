using System.Diagnostics;

namespace GLA
{
    public class Special : ArmyPart
    {
        [ExcelData(ColumnName = "Nom")]
        public string _Nom { get; set; }
        [ExcelData(ColumnName = "Description")]
        public string _Description { get; set; }

        public override void Resolve()
        {
        }
    }
}