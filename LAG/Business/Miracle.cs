using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace GLA
{
    [DebuggerDisplay("Miracle {Id} {_name}")]
    public class Miracle : ArmyPart
    {
        [ExcelData(ColumnName = "nom")]
        public string _name { get; set; }

        [ExcelData(ColumnName = "ferveur")]
        public string _fervor { get; set; }

        [ExcelData(ColumnName = "difficulte")]
        public string _difficulty { get; set; }

        [ExcelData(ColumnName = "culte")]
        public string _cult { get; set; }

        [ExcelData(ColumnName = "portee")]
        public string _range { get; set; }

        [ExcelData(ColumnName = "duree")]
        public string _duration { get; set; }

        [ExcelData(ColumnName = "aire")]
        public string _aoe { get; set; }

        [ExcelData(ColumnName = "foicreation")]
        public string _creation { get; set; }

        [ExcelData(ColumnName = "foialteration")]
        public string _alteration { get; set; }

        [ExcelData(ColumnName = "foidestruction")]
        public string _destruction { get; set; }

        [ExcelData(ColumnName = "effet")]
        public string _effect { get; set; }

        public List<string> Cults { get; set; }
       
        public override void Resolve()
        {
            Cults = _cult.Split(',').Select(v => v.Trim()).ToList();
        }
    }
}