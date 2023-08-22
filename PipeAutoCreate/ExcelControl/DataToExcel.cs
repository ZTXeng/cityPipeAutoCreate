using NPOI.SS.UserModel;
using PipeAutoCreate.DataModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.ExcelControl
{
    public static class DataToExcel
    {
        public static DataTable ConvertTo(List<Piping> pipings)
        {
            DataTable dt = new DataTable();

            dt = new DataTable(pipings.First().Profession);

            var titleNames = new List<string>();

            titleNames.Add("起点编号");
            titleNames.Add("终点编号");
            titleNames.Add("起点");
            titleNames.Add("终点");
            titleNames.Add("规格");

            var extens = pipings.FirstOrDefault().ExtensionParameters;
            foreach (var ext in extens)
            {
                titleNames.Add(ext.Name);
            }

            for (int i = 0; i < titleNames.Count; i++)
            {
                dt.Columns.Add(titleNames[i], typeof(string));
            }

            foreach (var pipe in pipings)
            {
                var dataRow = dt.NewRow();

                dataRow[0] = pipe.StartPoint.Id;
                dataRow[1] = pipe.EndPoint.Id;
                dataRow[2] = pipe.StartPoint.X.ToString() + ","
                            + pipe.StartPoint.Y.ToString() + ","
                            + pipe.StartPoint.Z.ToString();
                dataRow[3] = pipe.EndPoint.X.ToString() + ","
                            + pipe.EndPoint.Y.ToString() + ","
                            + pipe.EndPoint.Z.ToString();
                dataRow[4] = pipe.Specification;
                for (int i = 5; i < titleNames.Count; i++)
                {
                    dataRow[i] = pipe.ExtensionParameters[i-5].Value;
                }

                dt.Rows.Add(dataRow);
            }

            return dt;
        }

        public static DataTable ConvertTo(List<BuildingAttachment> attachments) {

            DataTable dt = new DataTable();

            dt = new DataTable(attachments.First().Profession);

            var titleNames = new List<string>();

            titleNames.Add("编号");
            titleNames.Add("坐标");
            titleNames.Add("规格");

            var extens = attachments.FirstOrDefault().ExtensionParameters;
            foreach (var ext in extens)
            {
                titleNames.Add(ext.Name);
            }

            for (int i = 0; i < titleNames.Count; i++)
            {
                dt.Columns.Add(titleNames[i], typeof(string));
            }


            foreach (var attachment in attachments)
            {
                var dataRow = dt.NewRow();

                dataRow[0] = attachment.Point.Id;
                dataRow[1] = attachment.Point.X.ToString() + ","
                            + attachment.Point.Y.ToString() + ","
                            + attachment.Point.Z.ToString();
                dataRow[2] = attachment.Specification;
                for (int i = 3; i < titleNames.Count; i++)
                {
                    dataRow[i] = attachment.ExtensionParameters[i - 3].Value;
                }

                dt.Rows.Add(dataRow);
            }

            return dt;
        }
    }
}
