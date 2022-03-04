using System;
using System.Collections.Generic;
using System.Reflection;

using Excel = Microsoft.Office.Interop.Excel;

namespace MS.ExcelData
{
   
    public class Column
    {
        public string Name { get; set; }
        public int Index { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="table"></param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        public Column(string name, Excel.ListObject table)
        {
            try
            {
                Name = name;
                Index = table.ListColumns[Name].Index;
            }
            catch (Exception)
            {
                throw new IndexOutOfRangeException($"Отсутствует столбец {Name} в таблице {table.Name}");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        /// <param name="table"></param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        public Column(int index, Excel.ListObject table)
        {
            try
            {
                Index = index;
                Name = table.ListColumns[1].Name;
            }
            catch (Exception)
            {
                throw new IndexOutOfRangeException($"Отсутствует столбец 1 в таблице {table.Name}");
            }
        }
    }

}
