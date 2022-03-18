using MS.ExcelData.Attributes;

using System;
using System.Collections.Generic;
using System.Reflection;

using Excel = Microsoft.Office.Interop.Excel;

namespace MS.ExcelData
{
    public class ExcelTable<TModel> : IExcelTable<TModel>
        where TModel : class, new()
    {
        private string _indexName;
        private readonly Dictionary<string, Column> _columns;
        private readonly Dictionary<string, PropertyInfo> _properties;
        private readonly Excel.Application _xlsApp;

        /// <summary>
        /// Таблица Excel с данными
        /// </summary>
        public Excel.ListObject Table { get; set; }

        /// <summary>
        /// Создание нового объекта данных
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="tableName"></param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        public ExcelTable(Excel.Workbook workbook) : this()
        {
            _xlsApp = workbook.Application;

            string tableName = GetTableName();
            SetTable(workbook, tableName);
            SetColumns();
        }

        public ExcelTable(Excel.Worksheet worksheet) : this()
        {
            _xlsApp = worksheet.Application;
            string tableName = GetTableName();
            SetTable(worksheet, tableName);
            SetColumns();
        }

        private ExcelTable()
        {
            _columns = new Dictionary<string, Column>();
            _properties = new Dictionary<string, PropertyInfo>();
            foreach (var property in typeof(TModel).GetProperties())
            {
                _properties.Add(property.Name, property);
            }
        }

        /// <summary>
        /// Сохранение данных в таблице
        /// </summary>
        /// <typeparam name="T">Тип объекта</typeparam>
        /// <param name="model">Экземпляр объекта</param>
        public virtual void Save(TModel model)
        {
            try
            {
                SetNonInteractive();
                int inx = FindRowIndexById(model);
                if (inx == 0)
                {
                    PropertyInfo property = GetIndexPropery(model);
                    if (property.PropertyType == typeof(int))
                    {
                        property.SetValue(model, (int)_xlsApp.WorksheetFunction.Max(Table.ListColumns[1].Range) + 1);
                    }

                    Create(model);
                }
                else
                {
                    Update(model);
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                _xlsApp.Interactive = true;
            }
        }

        /// <summary>
        /// Удаление записи из таблицы
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="model"></param>
        public virtual void Delete(TModel model)
        {
            try
            {
                SetNonInteractive();
                int inx = FindRowIndexById(model);
                Table.ListRows[inx].Delete();

            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                _xlsApp.Interactive = true;
            }
        }

        /// <summary>
        /// Получение списка с данными заданного типа
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public virtual IEnumerable<TModel> GetAll()
        {
            List<TModel> list = new List<TModel>();
            if (Table.DataBodyRange == null) return list;
            object[,] data = (object[,])Table.DataBodyRange.Value;

            for (int i = 1; i <= data.GetLength(0); i++)
            {
                Dictionary<int, object> dictionary = new Dictionary<int, object>();
                for (int j = 1; j <= data.GetLength(1); j++)
                {
                    dictionary.Add(j, data[i, j]);
                }

                TModel model = CreateModel(dictionary);
                list.Add(model);
            }

            return list;
        }

        /// <summary>
        /// Получает данные объекта по Id
        /// </summary>
        /// <typeparam name="T">Тип объекта</typeparam>
        /// <param name="keyValue">Идентификатор</param>
        /// <returns></returns>
        public virtual TModel GetById(object keyValue)
        {
            int inx = FindRowIndexByValueInColumn(keyValue, 1);
            if (inx == 0) return null;

            return CreateModel(GetDataFromRow(inx));
        }

        /// <summary>
        /// Получает данные по номеру строки
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="index"></param>
        /// <returns></returns>
        public virtual TModel GetByRowIndex(int index)
        {
            return CreateModel(GetDataFromRow(index));
        }

        /// <summary>
        /// Получение имени таблицы из атрибута
        /// </summary>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        private string GetTableName()
        {
            string tableName = string.Empty;
            foreach (Attribute attr in typeof(TModel).GetCustomAttributes(false))
            {
                if (attr is TableNameAttribute ageAttribute)
                {
                    tableName = ageAttribute.Name;
                    break;
                }
            }

            if (string.IsNullOrEmpty(tableName)) throw new ArgumentNullException("Отсутствует атрибут, указывающий название таблицы");

            return tableName;
        }

        /// <summary>
        /// Установка таблицы
        /// </summary>
        /// <param name="worksheet">Имя листа в котором расположена таблица</param>
        /// <param name="tableName">Имя таблицы</param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        private void SetTable(Excel.Worksheet worksheet, string tableName)
        {
            try
            {
                Table = worksheet.ListObjects[tableName];
            }
            catch (Exception)
            {
                throw new IndexOutOfRangeException($"Отсутствует таблица {tableName}"); ;
            }
        }

        /// <summary>
        /// Установка таблицы
        /// </summary>
        /// <param name="workbook">Книга, в которой находится таблица</param>
        /// <param name="tableName">Имя таблицы</param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        private void SetTable(Excel.Workbook workbook, string tableName)
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                foreach (Excel.ListObject item in worksheet.ListObjects)
                {
                    if (item.Name == tableName)
                    {
                        Table = item;
                        return;
                    }
                }
            }
            throw new IndexOutOfRangeException($"Отсутствует таблица {tableName}"); ;
        }

        /// <summary>
        /// Установка столбцов таблицы
        /// </summary>
        private void SetColumns()
        {
            foreach (var property in _properties)
            {
                Column column = new Column(Table);

                foreach (Attribute attr in property.Value.GetCustomAttributes(false))
                {
                    if (attr is NameAttribute columnName)
                    {
                        column.SetByName(columnName.Name);
                    }
                    if (attr is IndexAttribute columnIndex)
                    {
                        column.SetByIndex(columnIndex.Index);
                    }
                    if (attr is IsReadOnlyAttribute)
                    {
                        column.IsReadOnly = true;
                    }
                    if (attr is IsIndexAttribute)
                    {
                        column.IsIndex = true;
                    }
                }

                if (column.IsSetColumn)
                {
                    column.Type = property.Value.PropertyType;
                    _columns.Add(property.Key, column);

                    if (column.IsIndex) _indexName = property.Key;
                }
            }
        }

        /// <summary>
        /// Ожидание установки неинтерактивного режима
        /// </summary>
        private void SetNonInteractive()
        {
            while (_xlsApp.Interactive)
            {
                try
                {
                    _xlsApp.Interactive = false;
                }
                catch { }
            }
        }

        /// <summary>
        /// Создает новую строку в таблице и заполняет ее данными модели
        /// </summary>
        /// <param name="model">Модель</param>
        private void Create(TModel model)
        {
            Excel.ListRow row;
            try
            {
                row = Table.ListRows.AddEx();
            }
            catch
            {
                throw;
            }
            UpdateRowTable(model, row);
        }

        /// <summary>
        /// Обновление данных в таблице
        /// </summary>
        /// <param name="model"></param>
        private void Update(TModel model)
        {
            int inx = FindRowIndexById(model);
            UpdateRowTable(model, inx);
        }

        /// <summary>
        /// Обновление строки таблицы данными из модели
        /// </summary>
        /// <param name="model">модель данных</param>
        /// <param name="rowIndex">номер строки таблицы</param>
        private void UpdateRowTable(TModel model, int rowIndex)
        {
            Excel.ListRow row = Table.ListRows[rowIndex];
            UpdateRowTable(model, row);
        }

        /// <summary>
        /// Обновление строки таблицы данными из модели
        /// </summary>
        /// <param name="model">модель данных</param>
        /// <param name="row">строка таблицы Excel</param>
        private void UpdateRowTable(TModel model, Excel.ListRow row)
        {
            foreach (var item in GetDataFromModel(model))
            {
                Excel.Range rng = (Excel.Range)row.Range.Cells[1, item.Key];
                rng.Value = item.Value;
            }
        }

        /// <summary>
        /// Заполнение объекта данными из словаря
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private TModel CreateModel(Dictionary<int, object> data)
        {
            TModel model = new TModel();
            foreach (var column in _columns)
            {
                try
                {
                    Type type = column.Value.Type;

                    if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        type = type.GetGenericArguments()[0];
                        _properties[column.Key].SetValue(model, ToTypeNullable(type, data[column.Value.Index]));
                    }
                    else
                    {
                        _properties[column.Key].SetValue(model, ToType(type, data[column.Value.Index]));
                    }
                }
                catch (Exception ex)
                {
                    throw new ArgumentException($"Не удается записать значение {data[column.Value.Index]} " +
                        $"в столбец {column.Value.Name}\n {ex.Message}");
                }
            }
            return model;
        }

        /// <summary>
        /// Заполнение словаря данными
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="model"></param>
        /// <returns></returns>
        private Dictionary<int, object> GetDataFromModel(TModel model)
        {
            Dictionary<int, object> data = new Dictionary<int, object>();

            foreach (var item in _columns)
            {
                if (item.Value.IsReadOnly) continue;
                data.Add(item.Value.Index, _properties[item.Key].GetValue(model));
            }
            return data;
        }

        /// <summary>
        /// Получает данные строки в виде словаря (номер столбца, значение)
        /// </summary>
        /// <param name="indexRow"></param>
        /// <returns></returns>
        private Dictionary<int, object> GetDataFromRow(int indexRow)
        {
            Excel.ListRow row = Table.ListRows[indexRow];
            object[,] data = (object[,])row.Range.Value;

            Dictionary<int, object> dctionary = new Dictionary<int, object>();
            for (int i = 1; i <= data.Length; i++)
            {
                dctionary.Add(i, data[1, i]);
            }
            return dctionary;
        }

        /// <summary>
        /// Получает значение индексного свойства
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private object GetIndexValue(TModel model)
        {
            var property = GetIndexPropery(model);
            return property.GetValue(model);
        }

        /// <summary>
        /// Определяет индексное свойство
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException"></exception>
        private PropertyInfo GetIndexPropery(TModel model)
        {
            if (string.IsNullOrEmpty(_indexName) || !_properties.ContainsKey(_indexName))
            {
                throw new NullReferenceException("Отсутствует индексное свойство");
            }
            return _properties[_indexName];
        }

        /// <summary>
        /// Определяет индекс строки в которой расположены данные по индексному столбцу
        /// </summary>
        /// <param name="model">Объект данных</param>
        /// <returns></returns>
        private int FindRowIndexById(TModel model)
        {
            object value = GetIndexValue(model);
            return FindRowIndexByValueInColumn(value, 1);
        }

        /// <summary>
        /// Находит индекс строки по ключу и указанному номеру столбца
        /// </summary>
        /// <param name="keyValue">Значение ключа</param>
        /// <param name="indexColumn">Номер столбца для поиска</param>
        /// <returns></returns>
        private int FindRowIndexByValueInColumn(object keyValue, int indexColumn)
        {
            object[,] columnValue = Table.ListColumns[indexColumn].Range.Value2;

            int res = 0;
            foreach (var item in columnValue)
            {
                if (item.Equals(keyValue)) return res;
                res++;
            }
            return 0;
        }

        /// <summary>
        /// Преобразование данных в зависимости от типа
        /// </summary>
        /// <param name="type"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException"></exception>
        /// <exception cref="InvalidCastException"></exception>
        private object ToType(Type type, object value)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.String:
                    return Convert.ToString(value);
                case TypeCode.Double:
                    return Convert.ToDouble(value);
                case TypeCode.Decimal:
                    return Convert.ToDecimal(value);
                case TypeCode.DateTime:
                    try
                    {
                        return Convert.ToDateTime(value);
                    }
                    catch (Exception)
                    {
                        return DateTime.MinValue;
                    }
                case TypeCode.Boolean:
                    return Convert.ToBoolean(value);
                case TypeCode.Int32:
                    try
                    {
                        return Convert.ToInt32(value);
                    }
                    catch (Exception)
                    {
                        throw new ArgumentException($"Не удается преобразовать {value} в число");
                    }


                case TypeCode.Empty:
                    throw new NullReferenceException("The target type is null.");
                case TypeCode.Object:
                    throw new InvalidCastException(String.Format("Cannot convert {0}.", type.Name));
                case TypeCode.DBNull:
                    throw new NullReferenceException("The target type is null.");
                case TypeCode.Char:
                    return Convert.ToChar(value);
                case TypeCode.SByte:
                    return Convert.ToSByte(value);
                case TypeCode.Byte:
                    return Convert.ToByte(value);
                case TypeCode.Int16:
                    return Convert.ToInt16(value);
                case TypeCode.UInt16:
                    return Convert.ToUInt16(value);
                case TypeCode.UInt32:
                    return Convert.ToUInt32(value);
                case TypeCode.Int64:
                    return Convert.ToInt64(value);
                case TypeCode.UInt64:
                    return Convert.ToUInt64(value);
                case TypeCode.Single:
                    return Convert.ToSingle(value);
                default:
                    throw new InvalidCastException("Conversion not supported."); ;
            }
        }

        /// <summary>
        /// Преобразование данных в зависимости от типа
        /// </summary>
        /// <param name="type"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException"></exception>
        /// <exception cref="InvalidCastException"></exception>
        private object ToTypeNullable(Type type, object value)
        {
            if (value == null) return null;
            return ToType(type, value);
        }
    }

}
