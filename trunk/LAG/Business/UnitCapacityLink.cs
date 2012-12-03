using System;

namespace GLA
{
    public class UnitCapacityLink : ArmyPart
    {
        [ExcelData(ColumnName = "figurine_index")]
        public string _unitId { get; set; }

        [ExcelData(ColumnName = "figurine_capacite_index")]
        public string _capacityId { get; set; }

        public override void Resolve()
        {
            var referer = string.Format("Lien Figurine-Capacité ({0})", this.Id);
            var unit = new Reference<Unit>(Convert.ToInt32(_unitId));
            unit.ResolveReference(Army.Units, referer);
            var capacity = new Reference<Capacity>(Convert.ToInt32(_capacityId));
            capacity.ResolveReference(Army.Capacities, referer);
            if (unit.Entity != null && capacity.Entity != null)
            {
                if (capacity.Entity.AssociatedUnits.Contains(unit.Entity))
                    Warnings.Add("{0}: le lien entre {1} ({2}) et {3} ({4}) existe déjà, le doublon sera ignoré. (Il faudrait corriger la base).", referer, unit.Entity.FullName, unit.Entity.Id, capacity.Entity._name, capacity.Entity.Id);
                else
                    capacity.Entity.AssociatedUnits.Add(unit.Entity);
                if (unit.Entity.AssociatedCapacities.Contains(capacity.Entity))
                    Warnings.Add("{0}: le lien entre {1} ({2}) et {3} ({4}) existe déjà, le doublon sera ignoré. (Il faudrait corriger la base).", referer, unit.Entity.FullName, unit.Entity.Id, capacity.Entity._name, capacity.Entity.Id);
                else
                    unit.Entity.AssociatedCapacities.Add(capacity.Entity);
            }
        }
    }
}