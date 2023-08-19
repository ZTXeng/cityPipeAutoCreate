using Autodesk.Revit.Creation;
using PipeAutoCreate.ExcelControl;
using PipeAutoCreate.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.ViewModel
{
    public class ExcelShowViewModel:ViewModelBase<ExcelShowModel>
    {

        public ExcelShowViewModel(Document doc)
        {

        }

        public ExcelShowViewModel()
        {
            Model = new ExcelShowModel();

            Model.DataTable = ExcelToData.ReadExcel2DataTable(@"F:\西环管网数据-整理.xls", 0, true);
        }
    }
}
