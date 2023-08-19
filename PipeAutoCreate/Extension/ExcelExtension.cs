using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.Extension
{
    public static class ExcelExtension
    {
        public static double GetCellNumerica(ICell cell)
        {
            
            if (cell.CellType == CellType.Numeric)
            {
                return cell.NumericCellValue;
            }
            else if (cell.CellType == CellType.String)
            {
                return double.Parse(cell.StringCellValue);
            }

            return 0;
        }

        public static string GetCellString(ICell cell) {

            if (cell.CellType == CellType.Numeric)
            {
                return cell.NumericCellValue.ToString();
            }
            else if (cell.CellType == CellType.String)
            {
                return cell.StringCellValue;
            }
            return string.Empty;
        }
    }
}
