using System.Collections.Generic;

namespace MS.ExcelData
{
    public class BaseRepository<TModel> : IBaseRepository<TModel>
    {
        public IExcelTable<TModel> TableContext;

        public BaseRepository(IExcelTable<TModel> table)
        {
            TableContext = table;
        }

        public virtual void Save(TModel model) => TableContext.Save(model);
        public virtual void Delete(TModel model) => TableContext.Delete(model);
        public virtual IEnumerable<TModel> GetAll() => TableContext.GetAll();
        public virtual TModel GetById(object keyValue) => TableContext.GetById(keyValue);
        public virtual TModel GetByRowIndex(int index) => TableContext.GetByRowIndex(index);
    }
}
