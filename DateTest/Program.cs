using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PipeAutoCreate.DataModel;
using PipeAutoCreate.ExcelControl;
using PipeAutoCreate.Extension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DateTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var path = @"F:\西环管网数据-整理.xls";

            //var professionPipes = ReadPipe(path);

            //WritePipes(professionPipes);

            //var attachments = ExcelToData.ReadBuildingAttachment(path);
            //ExcelToData.WriteAttchments(attachments);
        }
    }
}

