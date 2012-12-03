using System.Linq;

namespace GLA
{
    public class FactionCombattantElement: ArmyPart
    {
        [ExcelData(ColumnName = "Faction")]
        public string _Faction { get; set; }
        [ExcelData(ColumnName = "Affili�s")]
        public string _Affilies { get; set; }
        [ExcelData(ColumnName = "Limit�s")]
        public string _Limites { get; set; }
        [ExcelData(ColumnName = "Interdits")]
        public string _Interdits { get; set; }

        public override void Resolve()
        {
            const string firstLineSyntaxMessage = "Feuille <Factions - combattants>, faction '{0}': la premi�re ligne doit contenir 'X;Y' dans la colonne <Limit�s> o� X est le co�t pour une troupe limit�e et Y le co�t pour un Champion limit�. Les cellules <Affili�s> et <Interdits> de cette ligne doivent �tre vides."; 

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
            //        Warnings.Add("Faction '{0}': la colonne <Affili�s> fait r�f�rence � '{1}' qui n'est pas trouv� dans la liste des unit�s (c'est peut-�tre normal).", _Faction, _Affilies);
            //}
            //if (!string.IsNullOrEmpty(_Limites))
            //{
            //    var limite = Army.Units.FirstOrDefault(u => u._Personnage == _Limites || u._Nom == _Limites);
            //    faction.Limites.Add(_Limites, limite);
            //    if (limite == null)
            //        Warnings.Add("Faction '{0}': la colonne <Limit�s> fait r�f�rence � '{1}' qui n'est pas trouv� dans la liste des unit�s (c'est peut-�tre normal).", _Faction, _Limites);
            //}
            //if (!string.IsNullOrEmpty(_Interdits))
            //{
            //    var interdit = Army.Units.FirstOrDefault(u => u._Personnage == _Interdits || u._Nom == _Interdits);
            //    faction.Interdits.Add(_Interdits, interdit);
            //    if (interdit == null)
            //        Warnings.Add("Faction '{0}': la colonne <Interdits> fait r�f�rence � '{1}' qui n'est pas trouv� dans la liste des unit�s (c'est peut-�tre normal).", _Faction, _Interdits);
            //}
        }
    }
}