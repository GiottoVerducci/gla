using System.Linq;

namespace GLA
{
    public class FactionCombattantElement: ArmyPart
    {
        [ExcelData(ColumnName = "Faction")]
        public string _Faction { get; set; }
        [ExcelData(ColumnName = "Affiliés")]
        public string _Affilies { get; set; }
        [ExcelData(ColumnName = "Limités")]
        public string _Limites { get; set; }
        [ExcelData(ColumnName = "Interdits")]
        public string _Interdits { get; set; }

        public override void Resolve()
        {
            const string firstLineSyntaxMessage = "Feuille <Factions - combattants>, faction '{0}': la première ligne doit contenir 'X;Y' dans la colonne <Limités> où X est le coût pour une troupe limitée et Y le coût pour un Champion limité. Les cellules <Affiliés> et <Interdits> de cette ligne doivent être vides."; 

            var faction = Army.Factions.FirstOrDefault(f => f.Nom == _Faction);
            if (faction == null)
            {
                faction = new Faction { Nom = _Faction };
                Army.Factions.Add(faction);
            }

            if (string.IsNullOrEmpty(_Affilies) && string.IsNullOrEmpty(_Interdits) && string.IsNullOrEmpty(faction.CoutLimite) && string.IsNullOrEmpty(faction.CoutLimiteChampion))
            {
                if(string.IsNullOrEmpty(_Limites))
                {
                    Warnings.Add(firstLineSyntaxMessage, _Faction);
                    return;
                }

                var couts = _Limites.Split(';').Select(c => c.Trim()).ToList();
                if (couts.Count != 2)
                {
                    Warnings.Add(firstLineSyntaxMessage, _Faction);
                    return;
                }

                faction.CoutLimite = couts[0];
                faction.CoutLimiteChampion = couts[1];
                return;
            }

            //if (!string.IsNullOrEmpty(_Affilies))
            //{
            //    var affilie = Army.Units.FirstOrDefault(u => u._Personnage == _Affilies || u._Nom == _Affilies);
            //    faction.Affilies.Add(_Affilies, affilie);
            //    if (affilie == null)
            //        Warnings.Add("Faction '{0}': la colonne <Affiliés> fait référence à '{1}' qui n'est pas trouvé dans la liste des unités (c'est peut-être normal).", _Faction, _Affilies);
            //}
            //if (!string.IsNullOrEmpty(_Limites))
            //{
            //    var limite = Army.Units.FirstOrDefault(u => u._Personnage == _Limites || u._Nom == _Limites);
            //    faction.Limites.Add(_Limites, limite);
            //    if (limite == null)
            //        Warnings.Add("Faction '{0}': la colonne <Limités> fait référence à '{1}' qui n'est pas trouvé dans la liste des unités (c'est peut-être normal).", _Faction, _Limites);
            //}
            //if (!string.IsNullOrEmpty(_Interdits))
            //{
            //    var interdit = Army.Units.FirstOrDefault(u => u._Personnage == _Interdits || u._Nom == _Interdits);
            //    faction.Interdits.Add(_Interdits, interdit);
            //    if (interdit == null)
            //        Warnings.Add("Faction '{0}': la colonne <Interdits> fait référence à '{1}' qui n'est pas trouvé dans la liste des unités (c'est peut-être normal).", _Faction, _Interdits);
            //}
        }
    }
}