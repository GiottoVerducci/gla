using System;

namespace GLA
{
    public class UnitFamilyLink : ArmyPart
    {
        [ExcelData(ColumnName = "figurine_index")]
        public string _unitId { get; set; }

        [ExcelData(ColumnName = "figurine_famille_index")]
        public string _familyId { get; set; }

        public override void Resolve()
        {
            var referer = string.Format("Lien Figurine-Famille ({0})", this.Id);
            var unit = new Reference<Unit>(Convert.ToInt32(_unitId));
            unit.ResolveReference(Army.Units, referer);
            var family = new Reference<Family>(Convert.ToInt32(_familyId));
            family.ResolveReference(Army.Families, referer);
            if (unit.Entity != null && family.Entity != null)
            {
                if (unit.Entity.AssociatedFamilies.Contains(family.Entity))
                    Warnings.Add("{0}: le lien entre {1} ({2}) et {3} ({4}) existe déjà, le doublon sera ignoré. (Il faudrait corriger la base).", referer, unit.Entity.FullName, unit.Entity.Id, family.Entity._name, family.Entity.Id);
                else
                    unit.Entity.AssociatedFamilies.Add(family.Entity);
            }
        }
    }
}