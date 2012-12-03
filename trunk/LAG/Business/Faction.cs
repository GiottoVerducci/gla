using System.Collections.Generic;
using System.Diagnostics;

namespace GLA
{
    [DebuggerDisplay("Faction {Nom}")]
    public class Faction
    {
        public string Nom { get; set; }
        public FactionElement Affiliation { get; set; }
        public List<FactionElement> Solos { get; private set; }
        public string CoutLimite { get; set; }
        public string CoutLimiteChampion { get; set; }

        public Dictionary<string, Unit> Affilies { get; private set; }
        public Dictionary<string, Unit> Limites { get; private set; }
        public Dictionary<string, Unit> Interdits { get; private set; }

        public Faction()
        {
            Solos = new List<FactionElement>();
            Affilies = new Dictionary<string, Unit>();
            Limites = new Dictionary<string, Unit>();
            Interdits = new Dictionary<string, Unit>();
        }
    }
}