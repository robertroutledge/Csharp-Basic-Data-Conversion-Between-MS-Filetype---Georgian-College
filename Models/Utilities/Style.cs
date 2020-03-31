using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace RoutledgeAssignmentFour.Models.Utilities
{
    public class Style
    {
        public static CellFormat cfBaseDate = new CellFormat()
        {
            ApplyNumberFormat = true,
            NumberFormatId = 14, 
        };
    }
}
