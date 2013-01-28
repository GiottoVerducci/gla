using System;

namespace GLA
{
    public class UnitMiracleLink : ArmyPart
    {
        [ExcelData(ColumnName = "figurine_index")]
        public string _unitId { get; set; }

        [ExcelData(ColumnName = "miracle_index")]
        public string _miracleId { get; set; }

        public override void Resolve()
        {
            var referer = "Lien Figurine-Miracle";
            var unit = new Reference<Unit>(Convert.ToInt32(_unitId));
            unit.ResolveReference(Army.Units, referer);
            var miracle = new Reference<Miracle>(Convert.ToInt32(_miracleId));
            miracle.ResolveReference(Army.Miracles, referer);
            if (unit.Entity != null && miracle.Entity != null)
            {
                if (unit.Entity.AssociatedMiracles.Contains(miracle.Entity))
                    Warnings.Add("{0}: le lien entre {1} ({2}) et {3} ({4}) existe déjà, le doublon sera ignoré. (Il faudrait corriger la base).", referer, unit.Entity.FullName, unit.Entity.Id, miracle.Entity._name, miracle.Entity.Id);
                else
                    unit.Entity.AssociatedMiracles.Add(miracle.Entity);
            }
        }
    }
}