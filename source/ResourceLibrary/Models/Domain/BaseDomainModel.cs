
using System;
using System.Collections.Generic;
using System.Text;
using Contensive.BaseClasses;
using System.Reflection;

namespace Contensive.Addons.ResourceLibrary.Models.Domain {
    public abstract class BaseDomainModel {
        private static T loadRecord<T>(CPBaseClass cp, CPCSBaseClass cs) where T : BaseDomainModel {
            T instance = null;
            try {
                if (cs.OK()) {
                    Type instanceType = typeof(T);
                    instance = (T)Activator.CreateInstance(instanceType);
                    foreach (PropertyInfo resultProperty in instance.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public)) {
                        switch (resultProperty.Name.ToLower()) {
                            case "specialcasefield":
                                break;
                            case "sortorder":
                                string sortOrder = cs.GetText(resultProperty.Name);
                                if (string.IsNullOrEmpty(sortOrder)) {
                                    sortOrder = "9999";
                                }
                                resultProperty.SetValue(instance, sortOrder, null);
                                break;
                            default:
                                switch (resultProperty.PropertyType.Name) {
                                    case "Int32":
                                        resultProperty.SetValue(instance, cs.GetInteger(resultProperty.Name), null);
                                        break;
                                    case "Boolean":
                                        resultProperty.SetValue(instance, cs.GetBoolean(resultProperty.Name), null);
                                        break;
                                    case "DateTime":
                                        resultProperty.SetValue(instance, cs.GetDate(resultProperty.Name), null);
                                        break;
                                    case "Double":
                                        resultProperty.SetValue(instance, cs.GetNumber(resultProperty.Name), null);
                                        break;
                                    default:
                                        resultProperty.SetValue(instance, cs.GetText(resultProperty.Name), null);
                                        break;
                                }
                                break;
                        }
                    }
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return instance;
        }
        protected static List<T> createListFromSql<T>(CPBaseClass cp, string sql, int pageSize, int pageNumber) where T : BaseDomainModel {
            List<T> result = new List<T>();
            try {
                CPCSBaseClass cs = cp.CSNew();
                List<string> ignoreCacheNames = new List<string>();
                Type instanceType = typeof(T);
                if (cs.OpenSQL(sql, "", pageSize, pageNumber)) {
                    T instance;
                    do {
                        instance = loadRecord<T>(cp, cs);
                        if (instance != null) {
                            result.Add(instance);
                        }
                        cs.GoNext();
                    } while (cs.OK());
                }
                cs.Close();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
            return result;
        }
    }
}
