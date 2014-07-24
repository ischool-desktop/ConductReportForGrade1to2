using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConductReportForGrade1to2
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"];
            item1["報表"]["成績相關報表"]["ConductGradeReport(for Grade 1-2)"].Enable = false;
            item1["報表"]["成績相關報表"]["ConductGradeReport(for Grade 1-2)"].Click += delegate
            {
                new Reporter(K12.Presentation.NLDPanels.Student.SelectedSource).ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.ConductGradeReport權限)
                {
                    item1["報表"]["成績相關報表"]["ConductGradeReport(for Grade 1-2)"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["ConductGradeReport(for Grade 1-2)"].Enable = false;
                }
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.ConductGradeReport, "ConductGradeReport(for Grade 1-2)"));
        }
    }
}
