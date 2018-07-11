using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Models
{
    public enum HorizontalAlignment
    {
        General = 0,
        Left = 1,
        Center = 2,
        CenterContinuous = 3,
        Right = 4,
        Fill = 5,
        Distributed = 6,
        Justify = 7
    }

    public enum VerticalAlignment
    {
        Top = 0,
        Center = 1,
        Bottom = 2,
        Distributed = 3,
        Justify = 4
    }
}
