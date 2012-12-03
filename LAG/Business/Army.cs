using System.Collections.Generic;

namespace GLA
{
    public class Army
    {
        public Dictionary<int, Unit> Units { get; private set; }
        public List<Special> Specials { get; private set; }
        public Dictionary<int, Artifact> Artifacts { get; private set; }
        public List<Faction> Factions { get; private set; }
        public Dictionary<int, Spell> Spells { get; private set; }
        public Dictionary<int, Miracle> Miracles { get; private set; }
        public Dictionary<int, Capacity> Capacities { get; private set; }
        public static Dictionary<int, Size> Sizes { get; private set; }
        public static Dictionary<int, Peuple> Peuples { get; private set; }
        public static Dictionary<int, Family> Families { get; private set; }
        public static Dictionary<int, MagicPath> MagicPaths { get; private set; }
        public static Dictionary<int, Rank> Ranks { get; private set; }
        public static Army AllArmies { get; private set; }

        public List<UnitArtifactLink> UnitArtifactLinks { get; private set; }
        public List<UnitCapacityLink> UnitCapacityLinks { get; private set; }
        public List<UnitFamilyLink> UnitFamilyLinks { get; private set; }
        public List<UnitMiracleLink> UnitMiracleLinks { get; private set; }
        public List<UnitSpellLink> UnitSpellLinks { get; private set; }
        public List<PeupleMagicPathLink> PeupleMagicPathLinks { get; private set; }

        public List<FactionElement> FactionElements { get; private set; }
        public List<FactionCombattantElement> FactionCombattantElements { get; private set; }

        public Dictionary<string, string> GeneralProperties { get; private set; }

        static Army()
        {
            AllArmies = new Army();
            Sizes = new Dictionary<int, Size>();
            Families = new Dictionary<int, Family>();
            Peuples = new Dictionary<int, Peuple>();
            MagicPaths = new Dictionary<int, MagicPath>();
            Ranks = new Dictionary<int, Rank>();
        }

        public Army()
        {
            Units = new Dictionary<int, Unit>();
            Specials = new List<Special>();
            Artifacts = new Dictionary<int, Artifact>();
            Factions = new List<Faction>();
            Spells = new Dictionary<int, Spell>();
            Miracles = new Dictionary<int, Miracle>();
            Capacities = new Dictionary<int, Capacity>();

            UnitArtifactLinks = new List<UnitArtifactLink>();
            UnitCapacityLinks = new List<UnitCapacityLink>();
            UnitFamilyLinks = new List<UnitFamilyLink>();
            UnitMiracleLinks = new List<UnitMiracleLink>();
            UnitSpellLinks = new List<UnitSpellLink>();
            PeupleMagicPathLinks = new List<PeupleMagicPathLink>();

            FactionElements = new List<FactionElement>();
            FactionCombattantElements = new List<FactionCombattantElement>();
            GeneralProperties = new Dictionary<string, string>();
        }

        public void Resolve()
        {
            foreach (var unit in Units.Values)
                unit.Resolve();
            foreach (var special in Specials)
                special.Resolve();
            foreach (var artefact in Artifacts.Values)
                artefact.Resolve();
            foreach (var factionElement in FactionElements)
                factionElement.Resolve();
            foreach (var factionCombattantElement in FactionCombattantElements)
                factionCombattantElement.Resolve();
            foreach (var sortilege in Spells.Values)
                sortilege.Resolve();
            foreach (var miracle in Miracles.Values)
                miracle.Resolve();
        }
    }
}