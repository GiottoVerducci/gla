using System.Collections.Generic;

namespace GLA
{
    public static class Warnings
    {
        private static readonly List<string> _warnings = new List<string>();

        public static void Clear()
        {
            _warnings.Clear();
        }

        public static void Add(string message, params object[] args)
        {
            _warnings.Add(string.Format(message, args));
        }

        public static IEnumerable<string> GetSummary()
        {
            return _warnings;
        }
    }
}