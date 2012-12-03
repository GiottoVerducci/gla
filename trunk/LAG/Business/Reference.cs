using System;
using System.Collections.Generic;
using GLA.Helper;

namespace GLA
{
    public class Reference<T>
        where T : class
    {
        private bool _isResolved;
        private T _entity;

        public int Id { get; private set; }

        public T Entity
        {
            get
            {
                if (!_isResolved)
                    throw new Exception("Not resolved!");
                return _entity;
            }
            private set { _entity = value; }
        }

        public Reference(int id)
        {
            this.Id = id;
        }

        public void ResolveReference(Dictionary<int, T> dictionary, string referer)
        {
            if (_isResolved)
                return;
            var entity = dictionary.ValueOrDefault(Id);
            if (entity == null)
                Warnings.Add("{0} : fait référence à {1} d'index {2} qui est introuvable. L'association sera ignorée.", referer, typeof(T), this.Id);
            else
                this.Entity = entity;
            _isResolved = true;
        }
    }
}