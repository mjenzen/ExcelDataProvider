using System;
using System.Collections.Generic;
using System.Data.Entity.Design.PluralizationServices;
using System.Dynamic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Excel.Data
{
    public class ExcelDataProvider : IDisposable
    {
        private readonly MsExcel.Application application;

        private readonly MsExcel.Workbook dataSource;

        protected ExcelDataProvider(string filePath)
        {
            application = new MsExcel.Application() { Visible = true };
            dataSource = application.Workbooks.Open(filePath);
        }

        protected IEnumerable<TEntity> GetObject<TEntity>() where TEntity : new()
        {
            var pluralizationService = PluralizationService.CreateService(System.Globalization.CultureInfo.CurrentCulture);
            return GetObject<TEntity>(pluralizationService.Pluralize(typeof(TEntity).Name));
        }

        protected IEnumerable<TEntity> GetObject<TEntity>(string sheetName) where TEntity : new()
        {
            var objectSheet = dataSource.Sheets[sheetName];
            MsExcel.Range range = objectSheet.UsedRange;

            Dictionary<int, string> ValueMap = GetValueMap(range);

            for (int row = 2; row <= range.Rows.Count; row++)
            {
                IDictionary<string, object> dynamicObject = new ExpandoObject();

                foreach (var keyValue in ValueMap)
                {
                    object val = range.Cells[row, keyValue.Key].value2;
                    string value = val == null ? "" : val.ToString();

                    dynamicObject[keyValue.Value] = value;
                }

                yield return BuildObject<TEntity>(dynamicObject);
            }
        }

        private Dictionary<int, string> GetValueMap(MsExcel.Range range)
        {
            var result = new Dictionary<int, string>();

            for (int i = 1; i <= range.Columns.Count; i++)
            {
                result.Add(i, (string)range.Cells[1, i].Value2);
            }

            return result;
        }

        private TEntity BuildObject<TEntity>(IDictionary<string, object> dynamicObject) where TEntity : new()
        {
            var newObject = new TEntity();

            foreach (var property in newObject.GetType().GetProperties())
            {
                object value;
                if (dynamicObject.TryGetValue(property.Name, out value))
                {
                    var parameter = Expression.Parameter(value.GetType());
                    var typedValue = ConvertValue(value.ToString(), property.PropertyType);

                    property.SetValue(newObject, typedValue);
                }
            }

            return newObject;
        }

        private object ConvertValue(string value, Type propertyType)
        {
            if (propertyType == typeof(string))
                return value;

            if (propertyType == typeof(int))
                return int.Parse(value);

            if (propertyType == typeof(int?))
                return String.IsNullOrEmpty(value) ? null : (int?)int.Parse(value);

            if (propertyType == typeof(double))
                return double.Parse(value);

            if (propertyType == typeof(double?))
                return String.IsNullOrEmpty(value) ? null : (double?)double.Parse(value);

            if (propertyType == typeof(float))
                return float.Parse(value);

            if (propertyType == typeof(float?))
                return String.IsNullOrEmpty(value) ? null : (float?)float.Parse(value);

            if (propertyType == typeof(decimal))
                return decimal.Parse(value);

            if (propertyType == typeof(decimal?))
                return String.IsNullOrEmpty(value) ? null : (decimal?)decimal.Parse(value);

            if (propertyType == typeof(long))
                return long.Parse(value);

            if (propertyType == typeof(long?))
                return String.IsNullOrEmpty(value) ? null : (long?)long.Parse(value);

            if (propertyType == typeof(bool))
                return bool.Parse(value);

            throw new NotImplementedException("Unsupported Type " + propertyType.Name);
        }

        public void Dispose()
        {
            dataSource.Save();
            dataSource.Close();

            application.Quit();
        }
    }
}
