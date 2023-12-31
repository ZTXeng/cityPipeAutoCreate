﻿using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.DataModel
{
    public class DPoint
    {
        public string Id { get; set; }

        public double X { get; set; }

        public double Y { get; set; }

        public double Z { get; set; }
          
        public DPoint() { }

        public XYZ Converter()
        {
            return new XYZ(X, Y, Z);
        }
    }
}
