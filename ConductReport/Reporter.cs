using Aspose.Words;
using Aspose.Words.Tables;
using CourseGradeB;
using CourseGradeB.StuAdminExtendControls;
using FISCA.Data;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ConductReportForGrade1to2
{
    public partial class Reporter : BaseForm
    {
        private int _schoolYear, _semester;
        AccessHelper _A;
        QueryHelper _Q;
        List<string> _ids;
        Dictionary<string, Dictionary<string, List<string>>> _template;
        BackgroundWorker _BW;

        public Reporter(List<string> ids)
        {
            InitializeComponent();
            _A = new AccessHelper();
            _Q = new QueryHelper();
            _ids = ids;
            _template = new Dictionary<string, Dictionary<string, List<string>>>();

            _BW = new BackgroundWorker();
            _BW.DoWork += new DoWorkEventHandler(BW_DoWork);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BW_Completed);

            _schoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            _semester = int.Parse(K12.Data.School.DefaultSemester);

            for (int i = -2; i <= 2; i++)
                cboSchoolYear.Items.Add(_schoolYear + i);

            cboSemester.Items.Add(1);
            cboSemester.Items.Add(2);

            cboSchoolYear.Text = _schoolYear + "";
            cboSemester.Text = _semester + "";

            LoadTemplate();
        }

        private void BW_DoWork(object sender, DoWorkEventArgs e)
        {
            string id = string.Join(",", _ids);

            //取得學生在指定學年度學期的科目老師對照
            string sqlcmd = "select sc_attend.id,sc_attend.ref_student_id,sc_attend.ref_course_id,teacher.teacher_name,tc_instruct.sequence,course.subject from sc_attend ";
            sqlcmd += "left join tc_instruct on tc_instruct.ref_course_id=sc_attend.ref_course_id ";
            sqlcmd += "left join teacher on teacher.id=tc_instruct.ref_teacher_id ";
            sqlcmd += "left join course on course.id=sc_attend.ref_course_id ";
            sqlcmd += "where ref_student_id in (" + id + ") and course.school_year=" + _schoolYear + " and course.semester=" + _semester;

            Dictionary<string, StudentObj> student_subj_teacher = new Dictionary<string, StudentObj>();
            DataTable dt = _Q.Select(sqlcmd);
            foreach (DataRow row in dt.Rows)
            {
                string student_id = row["ref_student_id"] + "";
                string teacher_name = row["teacher_name"] + "";
                string subject_name = row["subject"] + "";
                string key = student_id + "_" + subject_name;

                int i = 0;
                int sequence = int.TryParse(row["sequence"] + "", out i) ? i : 4;

                if (!student_subj_teacher.ContainsKey(key))
                {
                    student_subj_teacher.Add(key, new StudentObj(row));
                }

                if (sequence > student_subj_teacher[key].Sequence)
                    student_subj_teacher[key].TeacherName = teacher_name;
            }

            //取得指定學生conduct record
            List<ConductRecord> records = _A.Select<ConductRecord>("ref_student_id in (" + id + ") and school_year=" + _schoolYear + " and semester=" + _semester + " and not term is null");

            Dictionary<string, ConductObj> student_conduct = new Dictionary<string, ConductObj>();
            foreach (ConductRecord record in records)
            {
                string student_id = record.RefStudentId + "";
                if (!student_conduct.ContainsKey(student_id))
                    student_conduct.Add(student_id, new ConductObj(record));

                student_conduct[student_id].LoadRecord(record);
            }

            //科目顯示順序
            List<string> subject_order = _template.Keys.ToList();
            subject_order.Sort(Tool.GetSubjectCompare());

            //Group顯示順序
            List<string> group_order = new List<string>();

            foreach (string subj in subject_order)
            {
                foreach (string group in _template[subj].Keys)
                {
                    if (!group_order.Contains(group))
                        group_order.Add(group);
                }
            }
            group_order.Sort(delegate(string x, string y)
            {
                string xx = x.PadLeft(40, '0');
                string yy = y.PadLeft(40, '0');
                return xx.CompareTo(yy);
            });

            //取得缺席天數
            DataTable absence_dt = _Q.Select("select id from _udt_table where name='ischool.elementaryabsence'");
            if (absence_dt.Rows.Count > 0)
            {
                string str = string.Format("select ref_student_id,personal_days,sick_days from $ischool.elementaryabsence where ref_student_id in ({0}) and school_year={1} and semester={2}", id, _schoolYear, _semester);
                absence_dt = _Q.Select(str);

                foreach (DataRow row in absence_dt.Rows)
                {
                    string sid = row["ref_student_id"] + "";
                    string pd = row["personal_days"] + "";
                    string sd = row["sick_days"] + "";

                    if (student_conduct.ContainsKey(sid))
                    {
                        student_conduct[sid].PersonalDays = pd;
                        student_conduct[sid].SickDays = sd;
                    }
                }
            }

            //取得指定學生的班級導師
            Dictionary<string, string> student_class_teacher = new Dictionary<string, string>();
            foreach (SemesterHistoryRecord r in K12.Data.SemesterHistory.SelectByStudentIDs(_ids))
            {
                foreach (SemesterHistoryItem item in r.SemesterHistoryItems)
                {
                    if (item.SchoolYear == _schoolYear && item.Semester == _semester)
                    {
                        if (!student_class_teacher.ContainsKey(item.RefStudentID))
                            student_class_teacher.Add(item.RefStudentID, item.Teacher);

                        //上課天數
                        //if (student_conduct.ContainsKey(r.RefStudentID))
                        //{
                        //    student_conduct[r.RefStudentID].SchoolDays = item.SchoolDayCount + "";
                        //}
                    }
                }
            }

            //開始列印
            Document doc = new Document();

            //排序
            List<ConductObj> sortList = student_conduct.Values.ToList();

            sortList.Sort(delegate(ConductObj x, ConductObj y)
            {
                string xx = (x.Class.Name + "").PadLeft(20, '0');
                xx += (x.Student.SeatNo + "").PadLeft(10, '0');
                xx += (x.Student.Name + "").PadLeft(20, '0');

                string yy = (y.Class.Name + "").PadLeft(20, '0');
                yy += (y.Student.SeatNo + "").PadLeft(10, '0');
                yy += (y.Student.Name + "").PadLeft(20, '0');

                return xx.CompareTo(yy);
            });

            List<string> sortIDs = sortList.Select(x => x.Student.ID).ToList();

            //foreach (ConductObj obj in student_conduct.Values)
            foreach (string student_id in sortIDs)
            {
                //不應該會爆炸
                ConductObj obj = student_conduct[student_id];

                Dictionary<string, string> mergeDic = new Dictionary<string, string>();
                mergeDic.Add("姓名", obj.Student.Name + "(" + obj.Student.StudentNumber + ")");
                mergeDic.Add("班級", obj.Class.Name + "," + obj.Student.SeatNo);
                mergeDic.Add("學年度", (_schoolYear + 1911) + "-" + (_schoolYear + 1912));
                mergeDic.Add("學期", _semester + "");

                Document temp = new Aspose.Words.Document(new MemoryStream(Properties.Resources.temp));
                DocumentBuilder bu = new DocumentBuilder(temp);

                //bu.CellFormat.Borders.LineStyle = LineStyle.Double;
                bu.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                bu.MoveToMergeField("Conduct");
                //Table table = bu.StartTable();

                foreach (string subject in subject_order)
                {
                    Table table = bu.StartTable();
                    string teacherName = student_subj_teacher.ContainsKey(obj.StudentID + "_" + subject) ? student_subj_teacher[obj.StudentID + "_" + subject].TeacherName : "";

                    if (subject == "Homeroom")
                        teacherName = student_class_teacher.ContainsKey(obj.StudentID) ? student_class_teacher[obj.StudentID] : "";

                    bu.InsertCell();
                    table.AllowAutoFit = false;
                    bu.CellFormat.Width = 10;
                    bu.Write("Term1");
                    bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    bu.InsertCell();
                    bu.CellFormat.Width = 10;
                    bu.Write("Term2");
                    bu.InsertCell();
                    bu.CellFormat.Width = 100;
                    bu.Write(subject.PadRight(40, ' ') + "Teacher: " + teacherName.PadRight(20, ' '));
                    bu.EndRow();

                    foreach (string group in group_order)
                    {
                        if (!_template[subject].ContainsKey(group))
                            continue;

                        bu.InsertCell();
                        bu.CellFormat.Width = 10;
                        bu.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;

                        bu.InsertCell();
                        bu.CellFormat.Width = 10;
                        bu.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;

                        bu.InsertCell();
                        bu.CellFormat.Width = 100;
                        bu.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
                        bu.Write(group);
                        bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                        bu.EndRow();
                        
                        foreach (string title in _template[subject][group])
                        {
                            string key = subject + "_" + group + "_" + title;

                            bu.InsertCell();
                            bu.CellFormat.Width = 10;
                            string term1_ans = obj.term1.ContainsKey(key) ? obj.term1[key] : "";
                            bu.Write(term1_ans);
                            bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                            bu.InsertCell();
                            bu.CellFormat.Width = 10;
                            string term2_ans = obj.term2.ContainsKey(key) ? obj.term2[key] : "";
                            bu.Write(term2_ans);

                            bu.InsertCell();
                            bu.CellFormat.Width = 100;
                            bu.Write(title);
                            bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;

                            bu.EndRow();
                        }
                    }
                    bu.EndTable();
                    bu.Writeln();
                }

                //Homeroom Teacher's Comment:
                bu.InsertCell();
                bu.CellFormat.Width = 120;
                bu.Write("Homeroom Teacher's Comment:");
                bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                bu.EndRow();

                //Total Days of Absence
                bu.InsertCell();
                bu.CellFormat.Width = 120;
                bu.Write("Total Days of Absence: " + obj.GetTotalAbsence());
                bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                bu.EndRow();

                //Comment Term1 and Term2 Title
                bu.InsertCell();
                bu.CellFormat.Width = 60;
                bu.Write("Term1");
                bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                bu.InsertCell();
                bu.CellFormat.Width = 60;
                bu.Write("Term2");
                bu.EndRow();

                //Comment Term1 and Term2 Content
                bu.InsertCell();
                bu.CellFormat.Width = 60;
                bu.CellFormat.VerticalAlignment = CellVerticalAlignment.Top;
                bu.RowFormat.Height = 100;
                bu.Write(obj.Comment1 + "");
                bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;

                bu.InsertCell();
                bu.CellFormat.Width = 60;
                bu.CellFormat.VerticalAlignment = CellVerticalAlignment.Top;
                bu.Write(obj.Comment2 + "");
                bu.EndRow();
                bu.EndTable();

                temp.MailMerge.Execute(mergeDic.Keys.ToArray(), mergeDic.Values.ToArray());
                doc.Sections.Add(doc.ImportNode(temp.FirstSection, true));
            }

            doc.Sections.RemoveAt(0);

            e.Result = doc;
        }

        private void BW_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            Document doc = e.Result as Document;
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = "ConductGradeReport(for Grade 1-2).doc";
            save.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";

            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    doc.Save(save.FileName, Aspose.Words.SaveFormat.Doc);
                    System.Diagnostics.Process.Start(save.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }
        }

        private void LoadTemplate()
        {
            List<ConductSetting> list = _A.Select<ConductSetting>("grade=2");
            if (list.Count > 0)
            {
                ConductSetting setting = list[0];

                XmlDocument xdoc = new XmlDocument();
                if (!string.IsNullOrWhiteSpace(setting.Conduct))
                    xdoc.LoadXml(setting.Conduct);

                Dictionary<string, List<string>> extraItem = new Dictionary<string, List<string>>();

                //Add HRT item and get extra common item
                foreach (XmlElement elem in xdoc.SelectNodes("//Conduct[@Common]"))
                {
                    string group = elem.GetAttribute("Group");
                    bool common = elem.GetAttribute("Common") == "True" ? true : false;

                    if (!_template.ContainsKey("Homeroom"))
                        _template.Add("Homeroom", new Dictionary<string, List<string>>());

                    if (!_template["Homeroom"].ContainsKey(group))
                        _template["Homeroom"].Add(group, new List<string>());

                    if (common)
                    {
                        if (!extraItem.ContainsKey(group))
                            extraItem.Add(group, new List<string>());
                    }

                    foreach (XmlElement item in elem.SelectNodes("Item"))
                    {
                        string title = item.GetAttribute("Title");

                        if (!_template["Homeroom"][group].Contains(title))
                            _template["Homeroom"][group].Add(title);

                        if (common)
                        {
                            if (!extraItem[group].Contains(title))
                                extraItem[group].Add(title);
                        }
                    }
                }

                //Add Subject Item
                foreach (XmlElement elem in xdoc.SelectNodes("//Conduct[@Subject]"))
                {
                    string group = elem.GetAttribute("Group");
                    string subject = elem.GetAttribute("Subject");

                    if (!_template.ContainsKey(subject))
                        _template.Add(subject, new Dictionary<string, List<string>>());

                    if (!_template[subject].ContainsKey(group))
                        _template[subject].Add(group, new List<string>());

                    foreach (XmlElement item in elem.SelectNodes("Item"))
                    {
                        string title = item.GetAttribute("Title");

                        if (!_template[subject][group].Contains(title))
                            _template[subject][group].Add(title);
                    }

                    //add extra item to subject item
                    foreach (string extra_group in extraItem.Keys)
                    {
                        if (!_template[subject].ContainsKey(extra_group))
                            _template[subject].Add(extra_group, new List<string>());

                        foreach (string item in extraItem[extra_group])
                        {
                            if (!_template[subject][extra_group].Contains(item))
                                _template[subject][extra_group].Add(item);
                        }
                    }
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            _schoolYear = int.Parse(cboSchoolYear.Text);
            _semester = int.Parse(cboSemester.Text);

            if (_BW.IsBusy)
                MessageBox.Show("系統忙碌中,請稍後再試...");
            else
                _BW.RunWorkerAsync();
        }

        public class StudentObj
        {
            public string studentId, courseId, TeacherName, SubjectName;
            public int Sequence;

            public StudentObj(DataRow row)
            {
                studentId = row["ref_student_id"] + "";
                courseId = row["ref_course_id"] + "";
                TeacherName = row["teacher_name"] + "";
                SubjectName = row["subject"] + "";

                int i = 0;
                Sequence = int.TryParse(row["sequence"] + "", out i) ? i : 4;
            }
        }

        public class ConductObj
        {
            public static XmlDocument _xdoc;
            public Dictionary<string, string> term1 = new Dictionary<string, string>();
            public Dictionary<string, string> term2 = new Dictionary<string, string>();
            public string Comment1, Comment2;
            public string StudentID;
            public StudentRecord Student;
            public ClassRecord Class;
            public string PersonalDays, SickDays, SchoolDays;

            public ConductObj(ConductRecord record)
            {
                StudentID = record.RefStudentId + "";

                Student = K12.Data.Student.SelectByID(StudentID);
                Class = Student.Class;

                if (Student == null)
                    Student = new StudentRecord();

                if (Class == null)
                    Class = new ClassRecord();
            }

            public void LoadRecord(ConductRecord record)
            {
                string subj = record.Subject;
                if (string.IsNullOrWhiteSpace(subj))
                    subj = "Homeroom";

                string term = record.Term;

                //Comment
                if (subj == "Homeroom")
                {
                    if (term == "1")
                        Comment1 = record.Comment;

                    if (term == "2")
                        Comment2 = record.Comment;
                }
                
                //XML
                if (_xdoc == null)
                    _xdoc = new XmlDocument();

                _xdoc.RemoveAll();
                if (!string.IsNullOrWhiteSpace(record.Conduct))
                    _xdoc.LoadXml(record.Conduct);

                foreach (XmlElement elem in _xdoc.SelectNodes("//Conduct"))
                {
                    string group = elem.GetAttribute("Group");

                    foreach (XmlElement item in elem.SelectNodes("Item"))
                    {
                        string title = item.GetAttribute("Title");
                        string grade = item.GetAttribute("Grade");

                        if (term == "1")
                        {
                            if (!term1.ContainsKey(subj + "_" + group + "_" + title))
                                term1.Add(subj + "_" + group + "_" + title, grade);
                        }

                        if (term == "2")
                        {
                            if (!term2.ContainsKey(subj + "_" + group + "_" + title))
                                term2.Add(subj + "_" + group + "_" + title, grade);
                        }
                    }
                }
            }

            public int GetTotalAbsence()
            {
                int p = 0;
                int s = 0;
                int.TryParse(PersonalDays, out p);
                int.TryParse(SickDays, out s);

                return p + s;
            }
        }
    }
}
