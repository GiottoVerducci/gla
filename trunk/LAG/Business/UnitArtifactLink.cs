using System;

namespace GLA
{
    public class UnitArtifactLink : ArmyPart
    {
        [ExcelData(ColumnName = "figurine_index")]
        public string _unitId { get; set; }

        [ExcelData(ColumnName = "artefact_index")]
        public string _artefactId { get; set; }

        public override void Resolve()
        {
            var referer = "Lien Figurine-Artefact ({0})";
            var unit = new Reference<Unit>(Convert.ToInt32(_unitId));
            unit.ResolveReference(Army.Units, referer);
            var artefact = new Reference<Artifact>(Convert.ToInt32(_artefactId));
            artefact.ResolveReference(Army.Artifacts, referer);
            if (unit.Entity != null && artefact.Entity != null)
            {
                if (artefact.Entity.AssociatedUnits.Contains(unit.Entity))
                    Warnings.Add("{0}: le lien entre {1} ({2}) et {3} ({4}) existe déjà, le doublon sera ignoré. (Il faudrait corriger la base).", referer, unit.Entity.FullName, unit.Entity.Id, artefact.Entity._name, artefact.Entity.Id);
                else
                    artefact.Entity.AssociatedUnits.Add(unit.Entity);
                if(unit.Entity.AssociatedArtifacts.Contains(artefact.Entity))
                    Warnings.Add("{0}: le lien entre {1} ({2}) et {3} ({4}) existe déjà, le doublon sera ignoré. (Il faudrait corriger la base).", referer, unit.Entity.FullName, unit.Entity.Id, artefact.Entity._name, artefact.Entity.Id);
                else
                    unit.Entity.AssociatedArtifacts.Add(artefact.Entity);
            }
        }
    }

    public class PeupleMagicPathLink : ArmyPart
    {
        [ExcelData(ColumnName = "figurine_peuple_index")]
        public string _peupleId { get; set; }

        [ExcelData(ColumnName = "magie_voie_index")]
        public string _magicPathId { get; set; }

        public override void Resolve()
        {
            var referer = string.Format("Lien Peuple-Voie magie ({0})", this.Id);
            var peuple = new Reference<Peuple>(Convert.ToInt32(_peupleId));
            peuple.ResolveReference(Army.Peuples, referer);
            var magicPath = new Reference<MagicPath>(Convert.ToInt32(_magicPathId));
            magicPath.ResolveReference(Army.MagicPaths, referer);
            if (peuple.Entity != null && magicPath.Entity != null)
            {
                if (peuple.Entity.AssociatedMagicPaths.Contains(magicPath.Entity))
                    Warnings.Add("{0}: le lien entre {1} ({2}) et {3} ({4}) existe déjà, le doublon sera ignoré. (Il faudrait corriger la base).", referer, peuple.Entity._name, peuple.Entity.Id, magicPath.Entity._name, magicPath.Entity.Id);
                else
                    peuple.Entity.AssociatedMagicPaths.Add(magicPath.Entity);
            }
        }
    }
}