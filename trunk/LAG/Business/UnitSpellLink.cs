using System;

namespace GLA
{
    public class UnitSpellLink : ArmyPart
    {
        [ExcelData(ColumnName = "figurine_index")]
        public string _unitId { get; set; }

        [ExcelData(ColumnName = "magie_index")]
        public string _spellId { get; set; }

        public override void Resolve()
        {
            var referer = "Lien Figurine-Magie ({0})";
            var unit = new Reference<Unit>(Convert.ToInt32(_unitId));
            unit.ResolveReference(Army.Units, referer);
            var spell = new Reference<Spell>(Convert.ToInt32(_spellId));
            spell.ResolveReference(Army.Spells, referer);
            if (unit.Entity != null && spell.Entity != null)
            {
                if (unit.Entity.AssociatedSpells.Contains(spell.Entity))
                    Warnings.Add("{0}: le lien entre {1} ({2}) et {3} ({4}) existe déjà, le doublon sera ignoré. (Il faudrait corriger la base).", referer, unit.Entity.FullName, unit.Entity.Id, spell.Entity._name, spell.Entity.Id);
                else
                    unit.Entity.AssociatedSpells.Add(spell.Entity);
            }
        }
    }
}