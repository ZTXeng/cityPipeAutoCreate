using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.DataModel
{
    internal enum Profession
    {
        [Description("中国电信")]
        DX,

        [Description("供电")]
        GD,

        [Description("给水")]
        JS,

        [Description("路灯")]
        LD,

        [Description("中国联通")]
        LT,

        [Description("热力")]
        RL,

        [Description("天然气")]
        TR,

        [Description("污水")]
        WS,

        [Description("交通信号")]
        XH,

        [Description("中国移动")]
        YD,

        [Description("雨水")]
        YS,
    }
}
