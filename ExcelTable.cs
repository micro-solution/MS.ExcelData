using System;
using System.Collections.Generic;
using System.Reflection;

using Excel = Microsoft.Office.Interop.Excel;

namespace MS.ExcelData
{
    public abstract class ExcelTable<TModel> : IExcelTable<TModel>
        where TModel : class, new()
    {
        /// <summary>
        /// Таблица Excel с данными
        /// </summary>
        public Excel.ListObject Table { get; set; }

        /// <summary>
        /// Лист на котором находится таблица с данными
        /// </summary>
        public Excel.Worksheet Worksheet { get; set; }

        private PropertyInfo IndexProperty
        {
            get
            {
                foreach (PropertyInfo property in PropertyColumns)
                {
                    int inx = ((Column)property.GetValue(this)).Index;
                    if (inx == 1) _indexProperty = property;
                }
                return _indexProperty;
            }
        }
        private PropertyInfo _indexProperty;

        /// <summary>
        /// Свойства столбцов контекста данных
        /// </summary>
        private List<PropertyInfo> PropertyColumns
        {
            get
            {
                if (_propertyColumns == null)
                {
                    _propertyColumns = new List<PropertyInfo>();

                    foreach (var item in GetType().GetProperties())
                    {
                        if (item.PropertyType.Name != nameof(Column)) continue;
                        _propertyColumns.Add(item);
                    }
                }
                return _propertyColumns;
            }
        }

        private List<PropertyInfo> _propertyColumns;

        private readonly Excel.Application _app;

        /// <summary>
        /// Создание нового объекта данных
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="tableName"></param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        public ExcelTable(Excel.Worksheet worksheet, string tableName)
        {
            Worksheet = worksheet;

            _app = Worksheet.Application;
            try
            {
                Table = worksheet.ListObjects[tableName];
            }
            catch (Exception)
            {
                throw new IndexOutOfRangeException($"На листе {worksheet.Name} отсутствует таблица {tableName}"); ;
            }
        }

        /// <summary>
        /// Ожидание установки неинтерактивного режима
        /// </summary>
        private void SetNonInteractive()
        {
            while (_app.Interactive)
            {
                try
                {
                    _app.Interactive = false;
                }
                catch { }
            }
        }

        /// <summary>
        /// Сохранение данных в таблице
        /// </summary>
        /// <typeparam name="T">Тип объекта</typeparam>
        /// <param name="model">Экземпляр объекта</param>
        public void Save(TModel model)
        {
            try
            {
                SetNonInteractive();
                int inx = FindRowIndexById(model);
                if (inx == 0)
                {
                    Excel.Application app = Worksheet.Application;
                    PropertyInfo property = GetProperyIndex(model);
                    if (property.PropertyType == typeof(int))
                    {
                        property.SetValue(model, (int)app.WorksheetFunction.Max(Table.ListColumns[1].Range) + 1);
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
                _app.Interactive = true;
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
                _app.Interactive = true;
            }
        }


        /// <summary>
        /// Получение списка с данными заданного типа
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public List<TModel> GetAll()
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

                TModel model = Set(dictionary);
                list.Add(model);
            }

            return list;
        }

        /// <summary>
        /// Получает данные объекта из контекста
        /// </summary>
        /// <typeparam name="T">Тип данных</typeparam>
        /// <param name="keyColumn">Столбец расположения ключа</param>
        /// <param name="keyValue">Значение ключа</param>
        /// <returns></returns>
        public TModel GetByColumn(object keyValue, Column keyColumn)
        {
            int inx = FindRowIndexByValueInColumn(keyValue, keyColumn.Index);
            if (inx == 0) return null;

            return Set(GetDataFromRow(inx));
        }

        /// <summary>
        /// Получает данные объекта по Id
        /// </summary>
        /// <typeparam name="T">Тип объекта</typeparam>
        /// <param name="keyValue">Идентификатор</param>
        /// <returns></returns>
        public TModel GetById(object keyValue)
        {
            int inx = FindRowIndexByValueInColumn(keyValue, 1);
            if (inx == 0) return null;

            return Set(GetDataFromRow(inx));
        }

        /// <summary>
        /// Получает данные по номеру строки
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="index"></param>
        /// <returns></returns>
        public TModel GetByRowIndex(int index)
        {
            return Set(GetDataFromRow(index));
        }

        /// <summary>
        /// Создает новую строку в таблице и заполняет ее данными модели
        /// </summary>
        /// <typeparam name="T">Тип объекта</typeparam>
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
        /// <typeparam name="T"></typeparam>
        /// <param name="model"></param>
        private void Update(TModel model)
        {
            int inx = FindRowIndexById(model);
            UpdateRowTable(model, inx);
        }
        private void UpdateRowTable(TModel model, int rowIndex)
        {
            Excel.ListRow row = Table.ListRows[rowIndex];
            UpdateRowTable(model, row);
        }
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
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        private TModel Set(Dictionary<int, object> data)
        {
            TModel model = new TModel();
            PropertyInfo[] modelProperties = model.GetType().GetProperties();

            foreach (PropertyInfo property in PropertyColumns)
            {
                foreach (PropertyInfo propertyModel in modelProperties)
                {
                    if (property.Name == propertyModel.Name)
                    {
                        int inx = ((Column)property.GetValue(this)).Index;

                        Type type = propertyModel.PropertyType;

                        if (propertyModel.SetMethod != null)
                        {
                            try
                            {
                                propertyModel.SetValue(model, ToType(type, data[inx]));
                            }
                            catch (Exception ex)
                            {
                                throw new ArgumentException($"Не удается записать {property.Name}\n {ex.Message}");
                            }
                        }
                        break;
                    }
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
            PropertyInfo[] modelProperties = model.GetType().GetProperties();

            foreach (PropertyInfo property in PropertyColumns)
            {
                if (property.SetMethod == null) continue;

                foreach (PropertyInfo propertyModel in modelProperties)
                {
                    if (property.Name == propertyModel.Name)
                    {
                        int inx = ((Column)property.GetValue(this)).Index;
                        data.Add(inx, propertyModel.GetValue(model));
                        break;
                    }
                }
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
        /// <typeparam name="T"></typeparam>
        /// <param name="model"></param>
        /// <returns></returns>
        private object GetIndexValue(TModel model)
        {
            PropertyInfo property = GetProperyIndex(model);
            return property.GetValue(model);
        }
        /// <summary>
        /// Определяет индексное свойство
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="model"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException"></exception>
        private PropertyInfo GetProperyIndex(TModel model)
        {
            PropertyInfo[] modelProperties = model.GetType().GetProperties();
            foreach (PropertyInfo propertyModel in modelProperties)
            {
                if (IndexProperty.Name == propertyModel.Name)
                {
                    return propertyModel;
                }
            }
            throw new NullReferenceException("Отсутствует индексное свойство");
        }

        /// <summary>
        /// Определяет индекс строки в которой расположены данные по индексному столбцу
        /// </summary>
        /// <typeparam name="T">Тип данных</typeparam>
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
            Excel.Range column = Table.ListColumns[indexColumn].Range;

            try
            {
                double res = _app.WorksheetFunction.Match(keyValue, column, 0) - 1;
                return (int)res;
            }
            catch
            {
                return 0;
            }
            //Excel.Range find = column.Find(keyValue, LookAt: Excel.XlLookAt.xlWhole);

            //if (find == null) return 0;
            //return find.Row - Table.HeaderRowRange.Row;
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
    }

}
