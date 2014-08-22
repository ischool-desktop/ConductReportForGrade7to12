using Aspose.Words;
using CourseGradeB;
using CourseGradeB.StuAdminExtendControls;
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

namespace ConductReportForGrade7to12
{
    public partial class Reporter : BaseForm
    {
        private int _schoolYear, _semester;
        AccessHelper _A;
        List<string> _ids;
        Dictionary<string, List<string>> _hrt_template;
        BackgroundWorker _BW;
        string _校長, _主任;

        public Reporter(List<string> ids)
        {
            InitializeComponent();
            _A = new AccessHelper();
            _ids = ids;
            _hrt_template = new Dictionary<string, List<string>>();
            _校長 = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName").InnerText;
            _主任 = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("EduDirectorName").InnerText;

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

        private void LoadTemplate()
        {
            List<ConductSetting> list = _A.Select<ConductSetting>("grade=12");
            if (list.Count > 0)
            {
                ConductSetting setting = list[0];

                XmlDocument xdoc = new XmlDocument();
                if (!string.IsNullOrWhiteSpace(setting.Conduct))
                    xdoc.LoadXml(setting.Conduct);

                foreach (XmlElement elem in xdoc.SelectNodes("//Conduct[@Common]"))
                {
                    string group = elem.GetAttribute("Group");

                    foreach (XmlElement item in elem.SelectNodes("Item"))
                    {
                        string title = item.GetAttribute("Title");

                            if (!_hrt_template.ContainsKey(group))
                                _hrt_template.Add(group, new List<string>());

                            _hrt_template[group].Add(title);
                    }
                }
            }
        }

        private void BW_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            Document doc = e.Result as Document;
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = "ConductGradeReport(for Grade 7-12).doc";
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

        private void BW_DoWork(object sender, DoWorkEventArgs e)
        {
            string id = string.Join(",", _ids);

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
                    }
                }
            }

            //取得指定學生conduct record
            List<ConductRecord> records = _A.Select<ConductRecord>("ref_student_id in (" + id + ") and school_year=" + _schoolYear + " and semester=" + _semester + " and term is null and subject is null");

            Dictionary<string, ConductObj> student_conduct = new Dictionary<string, ConductObj>();
            foreach (ConductRecord record in records)
            {
                string student_id = record.RefStudentId + "";
                if (!student_conduct.ContainsKey(student_id))
                    student_conduct.Add(student_id, new ConductObj(record));

                student_conduct[student_id].LoadRecord(record);
            }

            //排序tempalte的item
            foreach (string group in _hrt_template.Keys)
                _hrt_template[group].Sort(delegate(string x, string y)
                {
                    return x.Length.CompareTo(y.Length);
                });

            //開始列印
            Document doc = new Document();

            foreach (ConductObj obj in student_conduct.Values)
            {
                Dictionary<string, string> mergeDic = new Dictionary<string, string>();
                mergeDic.Add("姓名", obj.Student.Name);
                mergeDic.Add("班級", obj.Class.Name);
                mergeDic.Add("座號", obj.Student.SeatNo + "");
                mergeDic.Add("學年度", (_schoolYear + 1911) + "-" + (_schoolYear + 1912));
                mergeDic.Add("學期", _semester == 1 ? _semester + "st" : _semester + "nd");
                mergeDic.Add("班導師", student_class_teacher.ContainsKey(obj.StudentID) ? student_class_teacher[obj.StudentID] : "");
                mergeDic.Add("校長", _校長);
                mergeDic.Add("主任", _主任);
                mergeDic.Add("Comment", obj.Comment);

                Document temp = new Aspose.Words.Document(new MemoryStream(Properties.Resources.template));
                DocumentBuilder bu = new DocumentBuilder(temp);

                bu.MoveToMergeField("Conduct");

                foreach (string group in _hrt_template.Keys)
                {
                    bu.Font.Bold = true;
                    bu.Font.Italic = true;
                    bu.Writeln(group);
                    
                    foreach(string item in _hrt_template[group])
                    {
                        string key = group + "_" + item;

                        string grade = obj.ConductGrade.ContainsKey(key) ? obj.ConductGrade[key] : "";

                        bu.Font.Bold = false;
                        bu.Font.Italic = false;
                        bu.Font.Color = Color.Red;
                        bu.Write(grade);

                        bu.Font.Color = Color.Black;
                        bu.Writeln(" " + item);
                    }
                }

                temp.MailMerge.Execute(mergeDic.Keys.ToArray(), mergeDic.Values.ToArray());
                doc.Sections.Add(doc.ImportNode(temp.FirstSection, true));
            }

            doc.Sections.RemoveAt(0);

            e.Result = doc;

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

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public class ConductObj
        {
            public static XmlDocument _xdoc;
            public Dictionary<string, string> ConductGrade = new Dictionary<string, string>();
            public string Comment;
            public string StudentID;
            public StudentRecord Student;
            public ClassRecord Class;

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
                //string subj = record.Subject;
                //if (string.IsNullOrWhiteSpace(subj))
                //    subj = "Homeroom";

                Comment = record.Comment;

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

                        if (!ConductGrade.ContainsKey(group + "_" + title))
                            ConductGrade.Add(group + "_" + title, grade);
                    }
                }
            }
        }
    }
}
