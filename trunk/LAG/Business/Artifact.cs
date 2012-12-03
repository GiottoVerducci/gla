using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using GLA.Helper;

namespace GLA
{
    [DebuggerDisplay("Artifact {Id} {_name}")]
    public class Artifact : ArmyPart
    {
        [ExcelData(ColumnName = "nom")]
        public string _name { get; set; }
        [ExcelData(ColumnName = "effet")]
        public string _effect { get; set; }
        [ExcelData(ColumnName = "npoints")]
        public string _points { get; set; }

        //public readonly List<Reference<Unit>> AssociatedUnits = new List<Reference<Unit>>();
        public readonly List<Unit> AssociatedUnits = new List<Unit>();

        private static string[] _prefixTexts = new[] { "L'", "Le", "La", "Les" };

        public string FullName
        {
            get
            {
                var index = _name.IndexOf('(');
                if (index == -1)
                    return _name;
                var length = _name.Substring(index).IndexOf(')') - 1;
                if (length <= 0)
                    return _name;
                var textBetweenParenthesis = _name.Substring(index + 1, length);
                if (_prefixTexts.Any(pt => string.Equals(pt, textBetweenParenthesis, StringComparison.InvariantCultureIgnoreCase)))
                    return string.Format("{0}{1}{2}", 
                        textBetweenParenthesis == "L'" ? "L'" : textBetweenParenthesis + " ", 
                        _name.Substring(0, index).Trim(),
                        _name.Substring(index + length + 2));
                return _name;
            }
        }

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
            //foreach (var unit in AssociatedUnits)
            //    unit.ResolveReference(Army.Units, string.Format("Artefact '{0}' ({1})", _name, Id));
        }
    }
}