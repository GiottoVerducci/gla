using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;

namespace GLA
{
    [DebuggerDisplay("Spell {Id} {_name}")]
    public class Spell : ArmyPart
    {
        [ExcelData(ColumnName = "nom")]
        public string _name { get; set; }
        [ExcelData(ColumnName = "puissance")]
        public string _power { get; set; }
        [ExcelData(ColumnName = "difficulte")]
        public string _difficulty { get; set; }

        [ExcelData(ColumnName = "voie")]
        public string _path { get; set; }

        [ExcelData(ColumnName = "portee")]
        public string _range { get; set; }

        [ExcelData(ColumnName = "duree")]
        public string _duration { get; set; }

        [ExcelData(ColumnName = "aire")]
        public string _aoe { get; set; }

        [ExcelData(ColumnName = "frequence")]
        public string _frequency { get; set; }

        [ExcelData(ColumnName = "manaprimagie")]
        public string _primagic { get; set; }


        [ExcelData(ColumnName = "manaair")]
        public string _air { get; set; }
        [ExcelData(ColumnName = "manaeau")]
        public string _water { get; set; }
        [ExcelData(ColumnName = "manafeu")]
        public string _fire { get; set; }
        [ExcelData(ColumnName = "manaterre")]
        public string _earth { get; set; }
        [ExcelData(ColumnName = "manalumiere")]
        public string _light { get; set; }
        [ExcelData(ColumnName = "manatenebre")]
        public string _darkness { get; set; }

        [ExcelData(ColumnName = "manainstinctive")]
        public string _instinct { get; set; }

        [ExcelData(ColumnName = "effet")]
        public string _effect { get; set; }

        public List<string> Voies { get; set; }

        public override void Resolve()
        {
            Voies = _path.Split(',').Select(v => v.Trim()).ToList();
        }
    }
}