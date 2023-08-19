using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.DataModel
{
    public class Piping
    {
        public DPoint StartPoint { get; set; }

        public DPoint EndPoint { get; set; }

        //专业
        public string Profession { set; get; }

        //特征
        public string Characteristic { get; set; }

        //规格
        public string Specification { get; set; }

        //套管尺寸
        public string CaseSize { get; set; }

        //管材
        public string Material { get; set; }

        //压力或电压
        public string Pressure { get; set; }

        //根数或流向
        public string PipeNumber { get; set; }

        //总孔数/已用孔数
        public string HoleNumber { get; set; }

        //埋设方式
        public string BuridType { get; set; }

        //建设日期
        public string ConstructionDate { get; set; }

        //所属道路
        public string AffiliatedRoad { get;set; }

        //备注
        public string Notes { get; set; }


        public Piping() { }

        public bool EquelPipe(Piping piping)
        {
            if (piping.StartPoint == null || piping.EndPoint == null
                || EndPoint == null)
            {
                return false;
            }
            
            else if (StartPoint.Id == piping.EndPoint.Id && EndPoint.Id == piping.StartPoint.Id)
            {
                return true;
            }
            else if (StartPoint.Id == piping.StartPoint.Id && EndPoint.Id == piping.EndPoint.Id)
            {
                return true;
            }
            return false;
        }
    }
}
