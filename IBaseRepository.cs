using System.Collections.Generic;

namespace MS.ExcelData
{
    public interface IBaseRepository<TModel>
    {
        void Save(TModel model);
        void Delete(TModel model);
        IEnumerable<TModel> GetAll();
        TModel GetById(object keyValue);
        TModel GetByRowIndex(int index);
    }
}
