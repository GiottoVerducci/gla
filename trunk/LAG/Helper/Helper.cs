using System.Collections.Generic;

namespace GLA.Helper
{
    public static class Helper
    {
        public static TV ValueOrDefault<TK, TV>(this Dictionary<TK, TV> dictionary, TK key)
        {
            TV value;
            dictionary.TryGetValue(key, out value);
            return value;
        }

        public static TV Ensure<TK, TV>(this Dictionary<TK, TV> dictionary, TK key, TV defaultValue)
        {
            TV value;
            if (!dictionary.TryGetValue(key, out value))
            {
                dictionary.Add(key, defaultValue);
                return defaultValue;
            }
            return value;
        }
    }
}
