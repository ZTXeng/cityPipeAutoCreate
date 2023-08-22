using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Math;
using PipeAutoCreate.DataModel;
using PipeAutoCreate.Extension;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PipeAutoCreate.ExcelControl
{
    public static class ExcelToData
    {
        public static List<Piping> ReadPipes(DataTable dataTable, List<string> popertyName)
        {
            var pipes = new List<Piping>();

            var rosCount = dataTable.Rows.Count;

            var titleStrings = GetFirstRow(dataTable);

            var startIdIndex = popertyName.FindIndex(x => x == "起点编号");
            var endIdIndex = popertyName.FindIndex(x => x == "终点编号");
            var Xindex = popertyName.FindIndex(x => x == "X");
            var Yindex = popertyName.FindIndex(x => x == "Y");
            var Zindex = popertyName.FindIndex(x => x == "管线高程");
            var speindex = popertyName.FindIndex(x => x == "规格");

            var startIds = GetStartIdIndexs(dataTable, startIdIndex);

            for (int i = 0; i < rosCount; i++)
            {
                var pipe = new Piping();

                var row = dataTable.Rows[i];

                var startId = row[startIdIndex] as string;
                var endId = row[endIdIndex] as string;
                var startX = double.Parse((row[Xindex] as string));
                var startY = double.Parse((row[Yindex] as string));
                var startZ = double.Parse((row[Zindex] as string));

                var endRowInstance = startIds.FirstOrDefault(x => x.Item1 == endId);
                if (endRowInstance.Equals(default((string,int)))) continue;

                var endIdRow = dataTable.Rows[startIds.First(x => x.Item1 == endId).Item2];

                var endX = double.Parse((endIdRow[Xindex] as string));
                var endY = double.Parse((endIdRow[Yindex] as string));
                var endZ = double.Parse((endIdRow[Zindex] as string));

                var startPoint = new DPoint() { Id = startId, X = startX, Y = startY, Z = startZ };
                var endPoint = new DPoint() { Id = endId, X = endX, Y = endY, Z = endZ };

                pipe.StartPoint = startPoint;
                pipe.EndPoint = endPoint;

                pipe.Specification = row[speindex].ToString();

                var columnCount = dataTable.Columns.Count;
                var subtractNumbers = SubtractNumbers(columnCount - 1, startIdIndex, endIdIndex, Xindex, Yindex, Zindex, speindex);
                foreach (var column in subtractNumbers)
                {
                    pipe.ExtensionParameters.Add(new ExtensionParameter() { Name = titleStrings[column], Value = row[column].ToString() });
                }

                pipe.Profession = dataTable.TableName;

                var repies = pipes.Where(x => (x.EquelPipe(pipe))).ToList();
                if (repies.Count() == 0 && pipe.EndPoint != null)
                {
                    pipes.Add(pipe);
                }
            }

            return pipes;
        }

        public static List<BuildingAttachment> ReadAttachments(DataTable dataTable, List<string> propertyName)
        {
            var attachments = new List<BuildingAttachment>();

            var rosCount = dataTable.Rows.Count;

            var titleStrings = GetFirstRow(dataTable);

            var startIdIndex = propertyName.FindIndex(x => x == "起点编号");
            var Xindex = propertyName.FindIndex(x => x == "X");
            var Yindex = propertyName.FindIndex(x => x == "Y");
            var Zindex = propertyName.FindIndex(x => x == "地面高程");
            var speindex = propertyName.FindIndex(x => x == "规格");
            var typenameindex = propertyName.FindIndex(x => x == "附属物");

            for (int i = 0; i < rosCount; i++)
            {
                var attachment = new BuildingAttachment();
                var row = dataTable.Rows[i];
                var startId = row[startIdIndex] as string;
                var startX = double.Parse((row[Xindex] as string));
                var startY = double.Parse((row[Yindex] as string));
                var startZ = double.Parse((row[Zindex] as string));
                var startPoint = new DPoint() { Id = startId, X = startX, Y = startY, Z = startZ };

                attachment.Point = startPoint;
                attachment.TypeName = row[typenameindex].ToString();
                attachment.Specification = row[speindex].ToString();

                var columnCount = dataTable.Columns.Count;
                var subtractNumbers = SubtractNumbers(columnCount - 1, startIdIndex, Xindex, Yindex, Zindex, speindex);
                foreach (var column in subtractNumbers)
                {
                    attachment.ExtensionParameters.Add(new ExtensionParameter() { Name = titleStrings[column], Value = row[column].ToString() });
                }

                attachment.Profession = dataTable.TableName;

                attachments.Add(attachment);
            }
            attachments  = attachments.DistinctBy(x => x.Point.Id).ToList();

            return attachments;
        }

        public static IWorkbook GetWorkBook(string path)
        {
            IWorkbook wk = null;
            var fs = File.OpenRead(path);

            if (path.Contains("xls"))
            {
                wk = new HSSFWorkbook(fs);
                //MessageBox.Show("若程序出错，请将EXCEL保存为2007版本以上");
                return wk;
            }
            else if (path.Contains("xlsx"))
            {
                wk = new XSSFWorkbook(fs);
                //MessageBox.Show("若程序出错，请将EXCEL保存为2007版本以下");
                return wk;
            }
            else
            {
                MessageBox.Show("请选择Excel文件");
                return null;
            }
        }

        public static List<(string, int)> GetStartIdIndexs(DataTable dataTable, int index)
        {
            var ids = new List<(string, int)>();
            var rowCount = dataTable.Rows.Count;
            for (int i = 1; i < rowCount; i++)
            {
                var row = dataTable.Rows[i];
                ids.Add((row[index].ToString(), i));
            }

            return ids;
        }

        public static void WritePipes(List<ProfessionPipe> pipings)
        {
            string basePath = System.Environment.CurrentDirectory;
            string resultPath = Path.Combine(basePath, "数据-反写.xls");
            var path = resultPath;

            var wk = new HSSFWorkbook();

            foreach (ProfessionPipe p in pipings)
            {
                var sheet = wk.CreateSheet(p.PProfession);
                for (int i = 0; i < p.Pipings.Count(); i++)
                {
                    var pipe = p.Pipings[i];
                    var row = sheet.CreateRow(i);
                    var cell = row.CreateCell(0);
                    cell.SetCellValue(pipe.StartPoint.Id);
                    var cell1 = row.CreateCell(1);
                    cell1.SetCellValue(pipe.EndPoint?.Id);
                    var cell2 = row.CreateCell(2);
                    cell2.SetCellValue(pipe.StartPoint.Z);
                }
            }

            using (var file = new FileStream(path, FileMode.Create))
            {
                wk.Write(file);
            }
        }

        public static void WriteAttchments(List<ProfessionAttachment> professionAttachments)
        {
            string basePath = System.Environment.CurrentDirectory;
            string resultPath = Path.Combine(basePath, "数据-反写2.xls");

            var path = resultPath;

            var wk = new HSSFWorkbook();

            foreach (ProfessionAttachment p in professionAttachments)
            {
                var sheet = wk.CreateSheet(p.Profession);
                for (int i = 0; i < p.BuildingAttachments.Count(); i++)
                {
                    var attachment = p.BuildingAttachments[i];
                    var row = sheet.CreateRow(i);
                    var cell = row.CreateCell(0);
                    cell.SetCellValue(attachment.Point.Id);
                    var cell1 = row.CreateCell(1);
                    cell1.SetCellValue(attachment.TypeName);
                    var cell2 = row.CreateCell(2);
                    cell2.SetCellValue(attachment.GroundHegiht);
                    var cell3 = row.CreateCell(3);
                    cell3.SetCellValue(attachment.Point.X);
                    var cell4 = row.CreateCell(4);
                    cell4.SetCellValue(attachment.Point.Y);
                }
            }

            using (var file = new FileStream(path, FileMode.Create))
            {
                wk.Write(file);
            }
        }

        public static List<DataTable> ReadExcelDataTbales(string path, bool isFirstRowHeader = true)
        {
            var dataTables = new List<DataTable>();
            var wk = GetWorkBook(path);

            var sheetCount = wk.NumberOfSheets;

            for (int i = 0; i < sheetCount; i++)
            {
                var dataTable = ReadExcel2DataTable(wk, i, isFirstRowHeader);
                dataTables.Add(dataTable);
            }

            return dataTables;
        }

        /// <summary>
        /// 从Excel读取数据并存入DataTable表中
        /// </summary>
        /// <param name="path">excel文件路径</param>
        /// <param name="sheetIndex">要读取的表index,从0开始</param>
        /// <param name="isFirstRowHeader">excel第一行数据是否作为列标题</param>
        /// <returns></returns>
        public static DataTable ReadExcel2DataTable(IWorkbook wk, int sheetIndex, bool isFirstRowHeader)
        {
            DataTable dt = new DataTable();

            ISheet sheet = wk.GetSheetAt(sheetIndex);  //读取索引的表数据

            dt = new DataTable(sheet.SheetName);   //dt.tableName

            //获取sheet的首行
            IRow headerRow = sheet.GetRow(0);
            int columnCount = headerRow.LastCellNum;
            int rowCount = sheet.LastRowNum;
            //dt新建列
            int startRowNum = 0; //当首行是标题时，数据从第1行开始读取。否则，从第0行读取。
            if (isFirstRowHeader)
            {
                for (int i = 0; i < columnCount; i++)
                {
                    dt.Columns.Add(headerRow.GetCell(i).ToString(), typeof(String));
                }
                startRowNum = 1;
            }
            else
            {
                for (int i = 0; i < columnCount; i++)
                {
                    dt.Columns.Add("列" + i, typeof(String));
                }
                startRowNum = 0;
            }

            //遍历

            for (int i = 0; i < rowCount; i++)  //
            {
                //每次开始遍历表时刷新列表
                DataRow dataRow = dt.NewRow();
                //dt.Rows.Add();
                IRow row = sheet.GetRow(i + startRowNum);  //从开始行
                if (row != null)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null)
                        {
                            //dt.Rows[i][j] = cell.ToString();
                            dataRow[j] = cell.ToString();
                        }
                    }

                    dt.Rows.Add(dataRow);
                }
            }

            return dt;
        }

        static List<int> SubtractNumbers(int num, params int[] subtracts)
        {
            List<int> result = new List<int>();

            for (int i = 0; i <= num; i++)
            {
                if (Array.IndexOf(subtracts, i) == -1)
                {
                    result.Add(i);
                }
            }

            return result;
        }

        public static List<string> GetFirstRow(DataTable dataTable)
        {
            var result = new List<string>();
            var columnCount = dataTable.Columns.Count;

            for (int i = 0; i < columnCount; i++)
            {
                result.Add(dataTable.Columns[i].ColumnName);
            }
            return result;
        }
    }
}
