using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.Extension
{
    public static class ListExtension
    {
        public static IEnumerable<T> DistinctBy<T>(this IEnumerable<T> source, Func<T, object> keySelector)
        {
            HashSet<object> uniqueKeys = new HashSet<object>();

            foreach (T item in source)
            {
                object key = keySelector(item);

                if (uniqueKeys.Add(key))
                {
                    yield return item;
                }
            }
        }
    }
}
