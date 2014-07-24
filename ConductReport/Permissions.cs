using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConductReportForGrade1to2
{
    class Permissions
    {
        public static string ConductGradeReport { get { return "ConductReport.6B287958-9E39-4342-A2B6-78DDA37724C0"; } }

        public static bool ConductGradeReport權限
        {
            get { return FISCA.Permission.UserAcl.Current[ConductGradeReport].Executable; }
        }
    }
}
