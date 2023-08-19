using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
        public static List<ProfessionAttachment> ReadBuildingAttachment(string path)
        {
            var professionAttachments = new List<ProfessionAttachment>();

            var wk = GetWorkBook(path);

            var sheetCount = wk.NumberOfSheets;

            for (int i = 0; i < sheetCount; i++)
            {
                var attachments = new List<BuildingAttachment>();
                var sheet = wk.GetSheetAt(i);

                var rowCount = sheet.LastRowNum;

                var ids = FindIdsInSheet(sheet);

                for (int j = 1; j < rowCount; j++)
                {
                    var buildingAttchment = new BuildingAttachment();

                    var rowDate = sheet.GetRow(j);

                    var id = rowDate.GetCell(0).StringCellValue;

                    var point = new DPoint()
                    {
                        Id = id,
                        X = ExcelExtension.GetCellNumerica(rowDate.GetCell(4)),
                        Y = ExcelExtension.GetCellNumerica(rowDate.GetCell(5)),
                        Z = ExcelExtension.GetCellNumerica(rowDate.GetCell(7))
                    };

                    buildingAttchment.Point = point;
                    buildingAttchment.Characteristic = ExcelExtension.GetCellString(rowDate.GetCell(2));
                    buildingAttchment.GroundHegiht = ExcelExtension.GetCellNumerica(rowDate.GetCell(6));
                    buildingAttchment.BuridHegiht = ExcelExtension.GetCellNumerica(rowDate.GetCell(8));

                    if (buildingAttchment.Characteristic == "上杆")
                    {
                        var reAttchemens = attachments.Where(x => x.Point.Id == id && x.TypeName == buildingAttchment.Characteristic);
                        if (reAttchemens.Count() == 0)
                        {
                            buildingAttchment.TypeName = buildingAttchment.Characteristic;
                            attachments.Add(buildingAttchment);
                        }
                    }

                    var tpyeName = ExcelExtension.GetCellString(rowDate.GetCell(3));

                    if (rowDate.GetCell(3).CellType != CellType.Blank)
                    {
                        var reAttchemens = attachments.Where(x => x.Point.Id == id && x.TypeName == tpyeName);
                        if (reAttchemens.Count() == 0)
                        {
                            buildingAttchment.TypeName = tpyeName;
                            attachments.Add(buildingAttchment);
                        }
                    }
                }

                professionAttachments.Add(new ProfessionAttachment()
                {
                    Profession = sheet.SheetName,
                    BuildingAttachments = attachments,
                });
            }

            return professionAttachments;
        }

        public static List<ProfessionPipe> ReadPipe(string path)
        {
            var professionPipes = new List<ProfessionPipe>();

            var wk = GetWorkBook(path);

            var sheetCount = wk.NumberOfSheets;

            for (int i = 0; i < sheetCount; i++)
            {
                var pipes = new List<Piping>();
                var sheet = wk.GetSheetAt(i);

                var rowCount = sheet.LastRowNum;

                var ids = FindIdsInSheet(sheet);

                for (int j = 1; j < rowCount; j++)
                {
                    var pipe = new Piping();

                    var rowDate = sheet.GetRow(j);
                    var va = rowDate.GetCell(0).StringCellValue;

                    var starstId = rowDate.GetCell(0);

                    var startId = rowDate.GetCell(0).StringCellValue;
                    var endId = rowDate.GetCell(1).StringCellValue;
                    var startPoint = new DPoint()
                    {
                        Id = startId,
                        X = ExcelExtension.GetCellNumerica(rowDate.GetCell(4)),
                        Y = ExcelExtension.GetCellNumerica(rowDate.GetCell(5)),
                        Z = ExcelExtension.GetCellNumerica(rowDate.GetCell(7))
                    };

                    var endIdIndexs = ids.Where(x => x.Item1 == endId);

                    if (endIdIndexs.Count() != 0)
                    {
                        var endIdIndex = endIdIndexs.First().Item2;
                        var endPoint = new DPoint()
                        {
                            Id = endId,
                            X = ExcelExtension.GetCellNumerica(sheet.GetRow(endIdIndex).GetCell(4)),
                            Y = ExcelExtension.GetCellNumerica(sheet.GetRow(endIdIndex).GetCell(5)),
                            Z = ExcelExtension.GetCellNumerica(sheet.GetRow(endIdIndex).GetCell(7)),
                        };

                        pipe.EndPoint = endPoint;
                    }

                    pipe.StartPoint = startPoint;

                    pipe.Characteristic = ExcelExtension.GetCellString(rowDate.GetCell(2));
                    pipe.Specification = ExcelExtension.GetCellString(rowDate.GetCell(9));
                    pipe.CaseSize = ExcelExtension.GetCellString(rowDate.GetCell(10));
                    pipe.Material = ExcelExtension.GetCellString(rowDate.GetCell(11));
                    pipe.Pressure = ExcelExtension.GetCellString(rowDate.GetCell(12));
                    pipe.PipeNumber = ExcelExtension.GetCellString(rowDate.GetCell(13));
                    pipe.HoleNumber = ExcelExtension.GetCellString(rowDate.GetCell(14));
                    pipe.BuridType = ExcelExtension.GetCellString(rowDate.GetCell(15));
                    pipe.ConstructionDate = ExcelExtension.GetCellString(rowDate.GetCell(16));
                    pipe.AffiliatedRoad = ExcelExtension.GetCellString(rowDate.GetCell(17));
                    pipe.Notes = ExcelExtension.GetCellString(rowDate.GetCell(18));
                    pipe.Profession = sheet.SheetName;

                    var repies = pipes.Where(x => (x.EquelPipe(pipe))).ToList();
                    if (repies.Count() == 0 && pipe.EndPoint != null)
                    {
                        pipes.Add(pipe);
                    }
                }

                professionPipes.Add(new ProfessionPipe()
                {
                    PProfession = sheet.SheetName,
                    Pipings = pipes
                });
            }

            return professionPipes;
        }

        public static IWorkbook GetWorkBook(string path)
        {
            IWorkbook wk = null;
            var fs = File.OpenRead(path);
            try
            {
                wk = new HSSFWorkbook(fs);
                return wk;
            }
            catch (Exception ex)
            {

                wk = new XSSFWorkbook(fs);
                return wk;
            }

        }

        public static List<(string, int)> FindIdsInSheet(ISheet sheet)
        {
            var ids = new List<(string, int)>();
            var rowCount = sheet.LastRowNum;
            for (int i = 1; i < rowCount; i++)
            {
                ids.Add((ExcelExtension.GetCellString(sheet.GetRow(i).GetCell(0)), i));
            }

            return ids;
        }

        public static void WritePipes(List<ProfessionPipe> pipings)
        {
            var path = @"F:\西环管网数据-反写.xls";

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
            var path = @"F:\西环管网数据-反写2.xls";

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
        /// <summary>
        /// 从Excel读取数据并存入DataTable表中
        /// </summary>
        /// <param name="path">excel文件路径</param>
        /// <param name="sheetIndex">要读取的表index,从0开始</param>
        /// <param name="isFirstRowHeader">excel第一行数据是否作为列标题</param>
        /// <returns></returns>
        public static DataTable ReadExcel2DataTable(string path, int sheetIndex, bool isFirstRowHeader)
        {
            //
            DataTable dt = new DataTable();
            FileStream fs = File.OpenRead(@path);  //打开EXCEL文件
            {
                IWorkbook wk = null;
                if (path.Contains("xls"))
                {
                    wk = new HSSFWorkbook(fs);      //把文件信息写入wk 
                                                    //MessageBox.Show("若程序出错，请将EXCEL保存为2007版本以上");
                }
                else if (path.Contains("xlsx"))
                {
                    wk = new XSSFWorkbook(fs);
                    //MessageBox.Show("若程序出错，请将EXCEL保存为2007版本以下");
                }
                else
                {
                    MessageBox.Show("请选择Excel文件");
                }
                //---------------------
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
            }
            return dt;
        }
    }
}
