using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace Uzzal.ExcelMapper
{
    public class ExcelMap<T>
    {
        private readonly Type _type;
        private readonly PropertyInfo[] _props;
        private readonly ExcelService _excelService;
        private readonly List<string> _columns = new List<string>();

        public ExcelMap(ExcelService service)
        {
            _excelService = service;
            _type = typeof(T);
            _props = _type.GetProperties();            
        }

        private void InitColumns()
        {            
            foreach (var item in _props)
            {
                var attr = item.GetCustomAttribute<ExcelColumnAttribute>();

                if (attr == null)
                {
                    _columns.Add(item.Name);
                }
                else
                {
                    _columns.Add(attr.Name);
                }
            }            
        }
        
        public List<T> Map(bool treatFirstRowAsHeader = true)
        {
            InitColumns();
            var rows = _excelService.GetRows();
            
            if (treatFirstRowAsHeader)
            {
                rows.RemoveAt(0);
            }

            var list = new List<T>();
            foreach (DataRow row in rows)
            {
                T obj = (T)Activator.CreateInstance(_type);
                SetProperty(obj, row);
                list.Add(obj);
            }

            return list;
        }

        private void SetProperty(T obj, DataRow row)
        {
            for (int i = 0; i < _props.Length; i++)
            {                
                try
                {
                    _props[i].SetValue(obj, row.Field<string>(_columns[i]));
                }
                catch (ArgumentException)
                {
                }
            }            
        }        
    }
}
