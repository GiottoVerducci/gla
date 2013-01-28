using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using GLA.Helper;

namespace GLA
{
    [DebuggerDisplay("{Id} {_name}")]
    public class Unit : ArmyPart
    {
        [ExcelData(ColumnName = "nom")]
        public string _name { get; set; }

        [ExcelData(ColumnName = "figurine_peuple_index")]
        public string _peupleId { get; set; }

        [ExcelData(ColumnName = "figurine_type_index")]
        public string _typeId { get; set; }

        [ExcelData(ColumnName = "figurine_rang_index")]
        public string _rankId { get; set; }

        [ExcelData(ColumnName = "npoints")]
        public string _points { get; set; }

        [ExcelData(ColumnName = "mouvement")]
        public string _M { get; set; }
        [ExcelData(ColumnName = "mouvementvol")]
        public string _alternateM { get; set; }
        [ExcelData(ColumnName = "initiative")]
        public string _INI { get; set; }
        [ExcelData(ColumnName = "attaque")]
        public string _ATT { get; set; }
        [ExcelData(ColumnName = "force")]
        public string _FOR { get; set; }
        [ExcelData(ColumnName = "defense")]
        public string _DEF { get; set; }
        [ExcelData(ColumnName = "resistance")]
        public string _RES { get; set; }
        [ExcelData(ColumnName = "tir")]
        public string _TIR { get; set; }
        [ExcelData(ColumnName = "courage")]
        public string _COU { get; set; }
        [ExcelData(ColumnName = "peur")]
        public string _PEU { get; set; }
        [ExcelData(ColumnName = "discipline")]
        public string _DIS { get; set; }

        [ExcelData(ColumnName = "pouvoir")]
        public string _POU { get; set; }
        [ExcelData(ColumnName = "foicreation")]
        public string _creation { get; set; }
        [ExcelData(ColumnName = "foialteration")]
        public string _alteration { get; set; }
        [ExcelData(ColumnName = "foidestruction")]
        public string _destruction { get; set; }

        [ExcelData(ColumnName = "foi")]
        public string _faith{ get; set; }

        [ExcelData(ColumnName = "aura")]
        public string _aura { get; set; }

        [ExcelData(ColumnName = "figurine_capacite_index")]
        public string _capacityId { get; set; }

        [ExcelData(ColumnName = "competence")]
        public string _competences { get; set; }

        [ExcelData(ColumnName = "equipement")]
        public string _equipment { get; set; }
        [ExcelData(ColumnName = "taille")]
        public string _size { get; set; }

        //	version

        [ExcelData(ColumnName = "concmouvement")]
        public string _concM { get; set; }
        [ExcelData(ColumnName = "concmouvementvol")]
        public string _concalternateM { get; set; }
        [ExcelData(ColumnName = "concinitiative")]
        public string _concINI { get; set; }
        [ExcelData(ColumnName = "concattaque")]
        public string _concATT { get; set; }
        [ExcelData(ColumnName = "concforce")]
        public string _concFOR { get; set; }
        [ExcelData(ColumnName = "concdefense")]
        public string _concDEF { get; set; }
        [ExcelData(ColumnName = "concresistance")]
        public string _concRES { get; set; }
        [ExcelData(ColumnName = "conctir")]
        public string _concTIR { get; set; }
        [ExcelData(ColumnName = "conccourage")]
        public string _concCOU { get; set; }
        [ExcelData(ColumnName = "concpeur")]
        public string _concPEU { get; set; }
        [ExcelData(ColumnName = "concdiscipline")]
        public string _concDIS { get; set; }

        [ExcelData(ColumnName = "figurine_genre_index")]
        public string _genreId { get; set; }

        public readonly List<Artifact> AssociatedArtifacts = new List<Artifact>();
        public readonly List<Capacity> AssociatedCapacities = new List<Capacity>();
        public readonly List<Miracle> AssociatedMiracles = new List<Miracle>();
        public readonly List<Spell> AssociatedSpells = new List<Spell>();
        public readonly List<Family> AssociatedFamilies = new List<Family>();

        public Reference<Size> Size { get; private set; }
        public Reference<Peuple> Peuple { get; private set; }
        public Reference<Rank> Rank { get; private set; }

        public bool IsPersonnage { get { return _typeId == "1"; } }
        public bool IsMagical { get { return new[] { 1, 11, 15, 21 }.Contains(Rank.Id); } }
        public bool IsDivine { get { return new[] { 9, 6, 7, 22 }.Contains(Rank.Id); } }

        public string FullName { get { return _name; } }
        public string Competences
        {
            get
            {
                if (string.IsNullOrEmpty(_competences))
                {
                    Warnings.Add("_competence is null or empty for {0} ({1})", FullName, Id);
                    _competences = "-";
                }
                var competences = _competences
                    .Split('(', ')')
                    .Select(t => t.Trim().Replace(".", null).Trim())
                    .Where(t => t.Length > 0)
                    .ToList();
                return string.Join(".\n", competences);
            }
        }
        
        public string Equipment
        {
            get { return _equipment == null ? string.Empty : _equipment.Trim().Replace(".", null).Trim(); }
        }

        public Unit ReferenceUnit { get; set; }
        public int VariationCount { get; set; }

        public static Dictionary<string, List<Unit>> GetGroupedUnits(List<Unit> associatedUnits)
        {
            var result = new Dictionary<string, List<Unit>>();
            for (int i = 0; i < associatedUnits.Count; i++)
            {
                var unit = associatedUnits[i];
                int p;
                if ((p = unit.FullName.LastIndexOf('(')) >= 0)
                {
                    var prefix = unit.FullName.Substring(0, p + 1);
                    var list = result.Ensure(prefix, new List<Unit>());
                    list.Add(unit);
                    while (i + 1 < associatedUnits.Count && associatedUnits[i + 1].FullName.LastIndexOf('(') >= 0 && associatedUnits[i + 1].FullName.StartsWith(prefix))
                    {
                        list.Add(associatedUnits[i + 1]);
                        ++i;
                    }
                }
                else
                {
                    result.Add(unit.FullName, new List<Unit> { unit });
                }
            }
            return result;
        }


//[ExcelData(ColumnName = "Spécial")]
        //public string _Special { get; set; }
        //[ExcelData(ColumnName = "Rang")]
        //public string _Rang { get; set; }
        //[ExcelData(ColumnName = "Peuple")]
        //public string _Peuple { get; set; }
        //[ExcelData(ColumnName = "M/F")]
        //public string _MF { get; set; }

        //public Special Special { get; set; }
        //public Taille Taille { get; set; }
        //public bool IsPersonnage { get { return !string.IsNullOrWhiteSpace(_Personnage); } }

        //public string Rang { get { return IsPersonnage ? "Champion " + _Rang.Substring(0, 1).ToLower() + _Rang.Substring(1) : _Rang; } }

        //public string FullName { get { return string.Join(", ", new[] { _Personnage, _name }.Where(s => !string.IsNullOrEmpty(s)).ToArray()); } }

        public override void Resolve()
        {
            Size = new Reference<Size>(Convert.ToInt32(_size));
            Size.ResolveReference(Army.Sizes, string.Format("Figurine {0} ({1})", _name, Id));
            Peuple = new Reference<Peuple>(Convert.ToInt32(_peupleId));
            Peuple.ResolveReference(Army.Peuples, string.Format("Figurine {0} ({1})", _name, Id) );
            Rank = new Reference<Rank>(Convert.ToInt32(_rankId));
            Rank.ResolveReference(Army.Ranks, string.Format("Figurine {0} ({1})", _name, Id));

            //Special = Army.Specials.FirstOrDefault(s => s._Nom == _Special);
            //Taille taille;
            //if (!Enum.TryParse(_Taille, true, out taille))
            //    Warnings.Add("Unité '{0}': taille '{1}' non reconnue.", _Taille);
            //else
            //    Taille = taille;
        }
    }
}