using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace GeoAddin
{
    internal class RemoveAllUnused: WipeOption
    {
        private readonly Document _doc;

       

        /// <inheritdoc/>
        internal override int Execute(string args = null)
        {
            var methods = new List<MethodInfo>
        {
            _doc.GetType().GetMethod("GetUnusedAppearances", BindingFlags.NonPublic | BindingFlags.Instance),
            _doc.GetType().GetMethod("GetUnusedMaterials", BindingFlags.NonPublic | BindingFlags.Instance),
            _doc.GetType().GetMethod("GetUnusedFamilies", BindingFlags.NonPublic | BindingFlags.Instance),
            _doc.GetType().GetMethod("GetUnusedImportCategories", BindingFlags.NonPublic | BindingFlags.Instance),
            _doc.GetType().GetMethod("GetUnusedStructures", BindingFlags.NonPublic | BindingFlags.Instance),
            _doc.GetType().GetMethod("GetUnusedSymbols", BindingFlags.NonPublic | BindingFlags.Instance),
            _doc.GetType().GetMethod("GetUnusedThermals", BindingFlags.NonPublic | BindingFlags.Instance),
            _doc.GetType().GetMethod("GetNonDeletableUnusedElements", BindingFlags.NonPublic | BindingFlags.Instance)
        };

            var num = 0;
            var tryCount = 0;
            while (true)
            {
                tryCount++;

                if (tryCount >= 5)
                    break;

                var hashSet = new HashSet<ElementId>();

                foreach (var methodInfo in methods)
                {
                    if (methodInfo?.Invoke(_doc, null) is ICollection<ElementId> c)
                    {
                        foreach (var id in c)
                            hashSet.Add(id);
                    }
                }

                if (hashSet.Count != num && hashSet.Count != 0)
                {
                    num += hashSet.Count;
                    using (var tr = new Transaction(_doc, "purge unused"))
                    {
                        tr.Start();
                        foreach (var elementId in hashSet)
                        {
                            try
                            {
                                _doc.Delete(elementId);
                            }
                            catch
                            {
                                num--;
                            }
                        }

                        tr.Commit();
                    }

                    continue;
                }

                break;
            }

           return num;
        }
    }
}
