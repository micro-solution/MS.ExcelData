using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;

namespace MS.ExcelData
{
    public interface IExcelTable<TModel>
    {
        Excel.ListObject Table { get; set; }
        Excel.Worksheet Worksheet { get; set; }

        List<TModel> GetAll();
        void Save(TModel model);
        void Delete(TModel model);
        TModel GetByColumn(object keyValue, Column keyColumn);
        TModel GetById(object keyValue);
        TModel GetByRowIndex(int index);

    }
}
