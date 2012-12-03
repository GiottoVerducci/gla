using System.Collections.Generic;
using System.Linq;

namespace GLA
{
    public class FactionElement: ArmyPart
    {
        [ExcelData(ColumnName = "Faction")]
        public string _Faction { get; set; }
        [ExcelData(ColumnName = "Nom")]
        public string _Nom { get; set; }
        [ExcelData(ColumnName = "Type")]
        public string _Type { get; set; }
        [ExcelData(ColumnName = "Réservé")]
        public string _Reserve { get; set; }
        [ExcelData(ColumnName = "PA")]
        public string _PA { get; set; }
        [ExcelData(ColumnName = "PA Champion")]
        public string _PAChampion { get; set; }
        [ExcelData(ColumnName = "Effet")]
        public string _Effet { get; set; }

        public Dictionary<string, Unit> Reserves { get; private set; }

        public FactionElement()
        {
            Reserves = new Dictionary<string, Unit>();
        }

        public override void Resolve()
        {
            //if (_Reserve != null)
            //{
            //    var reserves = _Reserve.Split(',').Select(r => r.Trim()).ToList();
            //    foreach (var reserve in reserves)
            //    {
            //        var unit = Army.Units.FirstOrDefault(u => u._Personnage == reserve || u._Nom == reserve);
            //        Reserves.Add(reserve, unit);

            //        if (unit == null)
            //        {
            //            Warnings.Add("Faction '{0}': la colonne <Réservé> de '{1}' fait référence à '{2}' qui n'est pas trouvé dans la liste des unités (c'est peut-être normal).", _Faction, _Nom, reserve);
            //        }
            //    }
            //}

            //var faction = Army.Factions.FirstOrDefault(f => f.Nom == _Faction);
            //if (faction == null)
            //{
            //    faction = new Faction { Nom = _Faction };
            //    Army.Factions.Add(faction);
            //}
            //switch (_Type)
            //{
            //    case "Affiliation":
            //        faction.Affiliation = this;
            //        break;
            //    case "Solo":
            //        faction.Solos.Add(this);
            //        break;
            //    default:
            //        Warnings.Add("Faction '{0}': type '{1}' inconnu (supporté: Affiliation ou Solo)", _Nom, _Type);
            //        break;
            //}
        }
    }
}