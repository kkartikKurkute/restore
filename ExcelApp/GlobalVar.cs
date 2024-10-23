using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApp
{
    
    public static class GlobalVar
    {
        public static Form3 frm3= new Form3();
        public static string GEmailId { get; set; }
        public static int GlobalUserId { get; set; }
        public static int CaseStudyId { get; set; }
        public static int QNo { get; set; }
        public static string strSheetName {get;set;}
        public static string Qhint { get; set; }
        public static string QLink { get; set; }
        public static int intClick { get; set; }
        public static int intOpen { get; set; }

    }
}
