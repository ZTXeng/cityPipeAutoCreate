using NPOI.SS.UserModel;
using PipeAutoCreate.DataModel;
using PipeAutoCreate.Extension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.ExcelControl
{
    public static class ExcelToOldData
    {
        //public static List<ProfessionAttachment> ReadBuildingAttachment(string path)
        //{
        //    var professionAttachments = new List<ProfessionAttachment>();

        //    var wk = GetWorkBook(path);

        //    var sheetCount = wk.NumberOfSheets;

        //    for (int i = 0; i < sheetCount; i++)
        //    {
        //        var attachments = new List<BuildingAttachment>();
        //        var sheet = wk.GetSheetAt(i);

        //        var rowCount = sheet.LastRowNum;

        //        var ids = FindIdsInSheet(sheet);

        //        for (int j = 1; j < rowCount; j++)
        //        {
        //            var buildingAttchment = new BuildingAttachment();

        //            var rowDate = sheet.GetRow(j);

        //            var id = rowDate.GetCell(0).StringCellValue;

        //            var point = new DPoint()
        //            {
        //                Id = id,
        //                X = ExcelExtension.GetCellNumerica(rowDate.GetCell(4)),
        //                Y = ExcelExtension.GetCellNumerica(rowDate.GetCell(5)),
        //                Z = ExcelExtension.GetCellNumerica(rowDate.GetCell(7))
        //            };

        //            buildingAttchment.Point = point;
        //            buildingAttchment.Characteristic = ExcelExtension.GetCellString(rowDate.GetCell(2));
        //            buildingAttchment.GroundHegiht = ExcelExtension.GetCellNumerica(rowDate.GetCell(6));
        //            buildingAttchment.BuridHegiht = ExcelExtension.GetCellNumerica(rowDate.GetCell(8));

        //            if (buildingAttchment.Characteristic == "上杆")
        //            {
        //                var reAttchemens = attachments.Where(x => x.Point.Id == id && x.TypeName == buildingAttchment.Characteristic);
        //                if (reAttchemens.Count() == 0)
        //                {
        //                    buildingAttchment.TypeName = buildingAttchment.Characteristic;
        //                    attachments.Add(buildingAttchment);
        //                }
        //            }

        //            var tpyeName = ExcelExtension.GetCellString(rowDate.GetCell(3));

        //            if (rowDate.GetCell(3).CellType != NPOI.SS.UserModel.CellType.Blank)
        //            {
        //                var reAttchemens = attachments.Where(x => x.Point.Id == id && x.TypeName == tpyeName);
        //                if (reAttchemens.Count() == 0)
        //                {
        //                    buildingAttchment.TypeName = tpyeName;
        //                    attachments.Add(buildingAttchment);
        //                }
        //            }
        //        }

        //        professionAttachments.Add(new ProfessionAttachment()
        //        {
        //            Profession = sheet.SheetName,
        //            BuildingAttachments = attachments,
        //        });
        //    }

        //    return professionAttachments;
        //}

        //public static List<ProfessionPipe> ReadPipe(string path)
        //{
        //    var professionPipes = new List<ProfessionPipe>();

        //    var wk = GetWorkBook(path);

        //    var sheetCount = wk.NumberOfSheets;

        //    for (int i = 0; i < sheetCount; i++)
        //    {
        //        var pipes = new List<Piping>();
        //        var sheet = wk.GetSheetAt(i);

        //        var rowCount = sheet.LastRowNum;

        //        var ids = FindIdsInSheet(sheet);



        //        for (int j = 1; j < rowCount; j++)
        //        {
        //            var pipe = new Piping();

        //            var rowDate = sheet.GetRow(j);
        //            var va = rowDate.GetCell(0).StringCellValue;

        //            var starstId = rowDate.GetCell(0);

        //            var startId = rowDate.GetCell(0).StringCellValue;
        //            var endId = rowDate.GetCell(1).StringCellValue;
        //            var startPoint = new DPoint()
        //            {
        //                Id = startId,
        //                X = ExcelExtension.GetCellNumerica(rowDate.GetCell(4)),
        //                Y = ExcelExtension.GetCellNumerica(rowDate.GetCell(5)),
        //                Z = ExcelExtension.GetCellNumerica(rowDate.GetCell(7))
        //            };

        //            var endIdIndexs = ids.Where(x => x.Item1 == endId);

        //            if (endIdIndexs.Count() != 0)
        //            {
        //                var endIdIndex = endIdIndexs.First().Item2;
        //                var endPoint = new DPoint()
        //                {
        //                    Id = endId,
        //                    X = ExcelExtension.GetCellNumerica(sheet.GetRow(endIdIndex).GetCell(4)),
        //                    Y = ExcelExtension.GetCellNumerica(sheet.GetRow(endIdIndex).GetCell(5)),
        //                    Z = ExcelExtension.GetCellNumerica(sheet.GetRow(endIdIndex).GetCell(7)),
        //                };

        //                pipe.EndPoint = endPoint;
        //            }

        //            pipe.StartPoint = startPoint;

        //            pipe.Characteristic = ExcelExtension.GetCellString(rowDate.GetCell(2));
        //            pipe.Specification = ExcelExtension.GetCellString(rowDate.GetCell(9));
        //            pipe.CaseSize = ExcelExtension.GetCellString(rowDate.GetCell(10));
        //            pipe.Material = ExcelExtension.GetCellString(rowDate.GetCell(11));
        //            pipe.Pressure = ExcelExtension.GetCellString(rowDate.GetCell(12));
        //            pipe.PipeNumber = ExcelExtension.GetCellString(rowDate.GetCell(13));
        //            pipe.HoleNumber = ExcelExtension.GetCellString(rowDate.GetCell(14));
        //            pipe.BuridType = ExcelExtension.GetCellString(rowDate.GetCell(15));
        //            pipe.ConstructionDate = ExcelExtension.GetCellString(rowDate.GetCell(16));
        //            pipe.AffiliatedRoad = ExcelExtension.GetCellString(rowDate.GetCell(17));
        //            pipe.Notes = ExcelExtension.GetCellString(rowDate.GetCell(18));
        //            pipe.Profession = sheet.SheetName;

        //            var repies = pipes.Where(x => (x.EquelPipe(pipe))).ToList();
        //            if (repies.Count() == 0 && pipe.EndPoint != null)
        //            {
        //                pipes.Add(pipe);
        //            }
        //        }

        //        professionPipes.Add(new ProfessionPipe()
        //        {
        //            PProfession = sheet.SheetName,
        //            Pipings = pipes
        //        });
        //    }

        //    return professionPipes;
        //}

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
    }
}
