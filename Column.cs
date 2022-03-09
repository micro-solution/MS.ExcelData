using System;

using Excel = Microsoft.Office.Interop.Excel;

namespace MS.ExcelData
{
    /// <summary>
    /// Класс столбца таблицы
    /// </summary>
    internal class Column
    {
        private readonly Excel.ListObject _table;

        /// <summary>
        /// Название столбца
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Номер столбца
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// Столбец только для чтения
        /// </summary>
        public bool IsReadOnly { get; set; } = false;

        /// <summary>
        /// Индексный столбец
        /// </summary>
        public bool IsIndex { get; set; } = false;

        /// <summary>
        /// Тип данных столбца
        /// </summary>
        public Type Type { get; set; }


        public bool IsSetColumn { get; private set; }

        public Excel.ListColumn ListColumn { get; private set; }

        public Column(Excel.ListObject table)
        {
            _table = table;
        }

        /// <summary>
        /// Установка столбца по имени
        /// </summary>
        /// <param name="name"></param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        public void SetByName(string name)
        {
            try
            {
                Name = name;
                ListColumn = _table.ListColumns[Name];
                Index = ListColumn.Index;
                IsSetColumn = true;
            }
            catch (Exception)
            {
                throw new IndexOutOfRangeException($"Отсутствует столбец {Name} в таблице {_table.Name}");
            }
        }

        /// <summary>
        /// Установка столбца по номеру
        /// </summary>
        /// <param name="index"></param>
        /// <param name="table"></param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        public void SetByIndex(int index)
        {
            try
            {
                Index = index;
                ListColumn = _table.ListColumns[index];
                Name = ListColumn.Name;
                IsSetColumn = true;
            }
            catch (Exception)
            {
                throw new IndexOutOfRangeException($"Отсутствует столбец 1 в таблице {_table.Name}");
            }
        }
    }
}
