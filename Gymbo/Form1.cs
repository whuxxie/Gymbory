using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.Spreadsheet;

namespace Gymbo
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private Hashtable NameListTable = new Hashtable();

        private static Worksheet worksheetschedule;
        private static Worksheet worksheetcount;
        private void btnopen_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Excel文件";
            ofd.FileName = "";
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            ofd.RestoreDirectory = true;
            ofd.RestoreDirectory = true;
            ofd.Filter = "Excel文件(*.xls)|*.xls|Excel2007文件(*.xlsx)|*.xlsx";
            ofd.ValidateNames = true; //文件有效性验证ValidateNames，验证用户输入是否是一个有效的Windows文件名
            ofd.CheckFileExists = true; //验证路径有效性
            ofd.CheckPathExists = true; //验证文件有效性
            string filePath = string.Empty;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;
            }

            if (!string.IsNullOrEmpty(filePath))
            {
                IWorkbook workbook = spreadsheetControl1.Document;
                workbook.LoadDocument(filePath);
            }

            xtc.SelectedTabPageIndex = 0;
        }

        private void btngenerate_Click(object sender, EventArgs e)
        {
            //Worksheet worksheet = spreadsheetControl1.ActiveWorksheet; 

            //IWorkbook workbook = spreadsheetControl1.Document;
            //workbook.LoadDocument(@"C:\Users\xx\Desktop\2017.10月课时2.xlsx");

            CreatHashTable();

            GenerateStatistics();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //FileStream fs=new FileStream("INSList.txt",FileMode.OpenOrCreate);

            //var s = "sdfklssjsd2234lfjsldfjljfjle343534535345jdsflg";
            //var maches = System.Text.RegularExpressions.Regex.Match(s, ".+\\d");
            //int iLastNumberIndex = maches.Length - 1;

            LoadNameList();
        }

        private void LoadNameList()
        {
            FileStream fs = new FileStream("INSList.txt", FileMode.OpenOrCreate);
            StreamReader sr = new StreamReader(fs);

            string line = sr.ReadLine().Trim();
            StringBuilder sb = new StringBuilder();
            while (!String.IsNullOrEmpty(line))
            {
                string[] names = line.Split(' ');
                NameListTable.Add(names[1], names[0]);

                sb.Append(line + "\r\n");
                line = sr.ReadLine();
            }
            meenamelist.Text = sb.ToString();
            sr.Close();
            fs.Close();
        }
        private void btnsavename_Click(object sender, EventArgs e)
        {
            try
            {
                FileStream fs = new FileStream("INSList.txt", FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);

                sw.Write(meenamelist.Text);
                sw.Close();
                fs.Close();

                MessageBox.Show("名单已保存");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
        }

        private void btnreflash_Click(object sender, EventArgs e)
        {
            LoadNameList();
        }

        static Hashtable nametable = new Hashtable();
        private void CreatHashTable()
        {
            nametable.Clear();
            string name = meenamelist.Text.Trim();

            string[] names = name.Split(new string[] { "\r\n" }, StringSplitOptions.None);

            foreach (DictionaryEntry entry in NameListTable)
            {
                int[] coursecounttable = new int[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                nametable.Add(entry.Key, coursecounttable);
            }
        }

        private void GenerateStatistics()
        {
            worksheetschedule = spreadsheetControl1.ActiveWorksheet;
            //string ss = worksheetschedule.Cells[3, 4].Value.TextValue;
            #region P
            ////Mon P
            //for (int i = 4; i < 36; i++)
            //{
            //    if (worksheetschedule.Cells[2, i].Value.TextValue.ToUpper() == "P")
            //    {
            //        for (int j = 3; j < 16; j++)
            //        {
            //            string value = worksheetschedule.Cells[j, i].Value.TextValue;
            //            if (value.Contains("(") || value.Contains(")") || value.Contains("（") || value.Contains("）") || value.Contains("-"))
            //            {
            //                value = value.Trim();
            //                //P7
            //                if (value.Trim().IndexOf("L7") >= 0)
            //                {
            //                    int start = value.IndexOf("(");
            //                    int end = value.IndexOf(")");
            //                    string teachers = value.Substring(start + 1, end - start - 1);
            //                    if (teachers.Contains('+'))
            //                    {
            //                        string[] teas = teachers.Split('+');
            //                        for (int k = 0; k < teas.Length; k++)
            //                        {
            //                            foreach (DictionaryEntry entry in nametable)
            //                            {
            //                                if (entry.Key.ToString().Substring(0, 2).ToUpper() == teas[k])
            //                                {
            //                                    int[] counttable = (int[])entry.Value;
            //                                    //P7主
            //                                    if (k == 0)
            //                                    {
            //                                        counttable[2] = counttable[2] + 1;
            //                                    }
            //                                    else//P7辅
            //                                    {
            //                                        counttable[3] = counttable[3] + 1;
            //                                    }
            //                                    nametable[entry.Key] = counttable;
            //                                }

            //                            }

            //                        }
            //                    }
            //                    else //只有一位老师
            //                    {
            //                        foreach (DictionaryEntry entry in nametable)
            //                        {
            //                            if (entry.Key.ToString().Substring(0, 2).ToUpper() == teachers)
            //                            {
            //                                int[] counttable = (int[])entry.Value;
            //                                counttable[2] = counttable[2] + 1;
            //                                nametable[entry.Key] = counttable;
            //                            }
            //                        }

            //                    }
            //                }
            //                else//其余play
            //                {
            //                    int start = value.IndexOf("(");
            //                    int end = value.IndexOf(")");
            //                    string teachers = value.Substring(start + 1, end - start - 1);
            //                    if (teachers.Contains('+'))
            //                    {
            //                        string[] teas = teachers.Split('+');
            //                        for (int k = 0; k < teas.Length; k++)
            //                        {
            //                            foreach (DictionaryEntry entry in nametable)
            //                            {
            //                                if (entry.Key.ToString().Substring(0, 2).ToUpper() == teas[k])
            //                                {
            //                                    int[] counttable = (int[])entry.Value;
            //                                    //P7主
            //                                    if (k == 0)
            //                                    {
            //                                        counttable[0] = counttable[0] + 1;
            //                                    }
            //                                    else//P7辅
            //                                    {
            //                                        counttable[1] = counttable[1] + 1;
            //                                    }
            //                                    nametable[entry.Key] = counttable;
            //                                }

            //                            }

            //                        }
            //                    }
            //                    else //只有一位老师
            //                    {
            //                        foreach (DictionaryEntry entry in nametable)
            //                        {
            //                            if (entry.Key.ToString().Substring(0, 2).ToUpper() == teachers)
            //                            {
            //                                int[] counttable = (int[])entry.Value;
            //                                counttable[0] = counttable[0] + 1;
            //                                nametable[entry.Key] = counttable;
            //                            }
            //                        }

            //                    }
            //                }
            //            }
            //        }
            //    }
            //#endregion
            //    #region M
            //    else if (worksheetschedule.Cells[2, i].Value.TextValue.ToUpper() == "M")
            //    {
            //        for (int j = 3; j < 16; j++)
            //        {
            //            string value = worksheetschedule.Cells[j, i].Value.TextValue;
            //            if (value.Contains("(") || value.Contains(")") || value.Contains("（") || value.Contains("）") || value.Contains("-"))
            //            {
            //                value = value.Trim();

            //                int start = value.IndexOf("(");
            //                int end = value.IndexOf(")");
            //                string teachers = value.Substring(start + 1, end - start - 1);
            //                if (teachers.Contains('+'))
            //                {
            //                    string[] teas = teachers.Split('+');
            //                    for (int k = 0; k < teas.Length; k++)
            //                    {
            //                        foreach (DictionaryEntry entry in nametable)
            //                        {
            //                            if (entry.Key.ToString().Substring(0, 2).ToUpper() == teas[k])
            //                            {
            //                                int[] counttable = (int[])entry.Value;
            //                                //P7主
            //                                if (k == 0)
            //                                {
            //                                    counttable[4] = counttable[4] + 1;
            //                                }
            //                                else//P7辅
            //                                {
            //                                    counttable[4] = counttable[4] + 1;
            //                                }
            //                                nametable[entry.Key] = counttable;
            //                            }

            //                        }

            //                    }
            //                }
            //                else //只有一位老师
            //                {
            //                    foreach (DictionaryEntry entry in nametable)
            //                    {
            //                        if (entry.Key.ToString().Substring(0, 2).ToUpper() == teachers)
            //                        {
            //                            int[] counttable = (int[])entry.Value;
            //                            counttable[4] = counttable[4] + 1;
            //                            nametable[entry.Key] = counttable;
            //                        }
            //                    }

            //                }
            //            }
            //        }

            //    }
            //    #endregion
            //    #region A
            //    else if (worksheetschedule.Cells[2, i].Value.TextValue.ToUpper() == "A")
            //    {
            //        for (int j = 3; j < 16; j++)
            //        {
            //            string value = worksheetschedule.Cells[j, i].Value.TextValue;
            //            if (value.Contains("(") || value.Contains(")") || value.Contains("（") || value.Contains("）") || value.Contains("-"))
            //            {
            //                value = value.Trim();
            //                //A3
            //                if (value.Trim().IndexOf("A3") >= 0)
            //                {
            //                    int start = value.IndexOf("(");
            //                    int end = value.IndexOf(")");
            //                    string teachers = value.Substring(start + 1, end - start - 1);
            //                    if (teachers.Contains('+'))
            //                    {
            //                        string[] teas = teachers.Split('+');
            //                        for (int k = 0; k < teas.Length; k++)
            //                        {
            //                            foreach (DictionaryEntry entry in nametable)
            //                            {
            //                                if (entry.Key.ToString().Substring(0, 2).ToUpper() == teas[k])
            //                                {
            //                                    int[] counttable = (int[])entry.Value;
            //                                    //P7主
            //                                    if (k == 0)
            //                                    {
            //                                        counttable[5] = counttable[5] + 1;
            //                                    }
            //                                    else//P7辅
            //                                    {
            //                                        counttable[6] = counttable[6] + 1;
            //                                    }
            //                                    nametable[entry.Key] = counttable;
            //                                }

            //                            }

            //                        }
            //                    }
            //                    else //只有一位老师
            //                    {
            //                        foreach (DictionaryEntry entry in nametable)
            //                        {
            //                            if (entry.Key.ToString().Substring(0, 2).ToUpper() == teachers)
            //                            {
            //                                int[] counttable = (int[])entry.Value;
            //                                counttable[5] = counttable[5] + 1;
            //                                nametable[entry.Key] = counttable;
            //                            }
            //                        }

            //                    }
            //                }
            //                else//其余Art
            //                {
            //                    int start = value.IndexOf("(");
            //                    int end = value.IndexOf(")");
            //                    string teachers = value.Substring(start + 1, end - start - 1);
            //                    if (teachers.Contains('+'))
            //                    {
            //                        string[] teas = teachers.Split('+');
            //                        for (int k = 0; k < teas.Length; k++)
            //                        {
            //                            foreach (DictionaryEntry entry in nametable)
            //                            {
            //                                if (entry.Key.ToString().Substring(0, 2).ToUpper() == teas[k])
            //                                {
            //                                    int[] counttable = (int[])entry.Value;
            //                                    //P7主
            //                                    if (k == 0)
            //                                    {
            //                                        counttable[7] = counttable[7] + 1;
            //                                    }
            //                                    else//P7辅
            //                                    {
            //                                        counttable[7] = counttable[7] + 1;
            //                                    }
            //                                    nametable[entry.Key] = counttable;
            //                                }

            //                            }

            //                        }
            //                    }
            //                    else //只有一位老师
            //                    {
            //                        foreach (DictionaryEntry entry in nametable)
            //                        {
            //                            if (entry.Key.ToString().Substring(0, 2).ToUpper() == teachers)
            //                            {
            //                                int[] counttable = (int[])entry.Value;
            //                                counttable[7] = counttable[7] + 1;
            //                                nametable[entry.Key] = counttable;
            //                            }
            //                        }

            //                    }
            //                }
            //            }
            //        }
            //    }
            //    #endregion
            //    #region OP
            //    else if (worksheetschedule.Cells[2, i].Value.TextValue.ToUpper() == "OP")
            //    {
            //        for (int j = 3; j < 16; j++)
            //        {
            //            string value = worksheetschedule.Cells[j, i].Value.TextValue;
            //            if (value.Contains("(") || value.Contains(")") || value.Contains("（") || value.Contains("）") || value.Contains("-"))
            //            {
            //                value = value.Trim();
            //                //HL

            //            }
            //        }
            //    }
            #endregion
            try
            {

                //}

                #region

                //for (int i = 1; i < 38; i++)
                //{
                //    for (int j = 3; j < 22; j++)
                //    {
                //        if (worksheetschedule.Cells[j, i].Value.TextValue == null)
                //        {
                //            continue;
                //        }
                //        string value = worksheetschedule.Cells[j, i].Value.TextValue.Trim();
                //        if (value.Contains("(") || value.Contains(")") || value.Contains("（") || value.Contains("）") ||
                //            value.Contains("-"))
                //        {
                //            int start = value.IndexOf("(") > -1 ? value.IndexOf("(") : value.IndexOf("（");
                //            int end = value.IndexOf(")") > -1 ? value.IndexOf(")") : value.IndexOf("）");
                //            string teacher = value.Substring(start + 1, end - start - 1);
                //            string[] teacherlist = teacher.Split('+');
                //            if (teacherlist.Length < 1)
                //            {
                //                return;
                //            }
                //            else
                //            {
                //                string symbol = value.Substring(0, 2);
                //                switch (symbol)
                //                {
                //                    case "L1":
                //                    case "L2":
                //                    case "L3":
                //                    case "L4":
                //                    case "L5":
                //                    case "L6":
                //                        for (int k = 0; k < teacherlist.Length; k++)
                //                        {
                //                            int[] counttable = new int[11];
                //                            string name = "";
                //                            foreach (DictionaryEntry entry in nametable)
                //                            {
                //                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                //                                {
                //                                    counttable = (int[])entry.Value;
                //                                    name = entry.Key.ToString();
                //                                    if (k == 0)
                //                                    {
                //                                        counttable[0] = counttable[0] + 1;
                //                                    }
                //                                    else
                //                                    {
                //                                        counttable[1] = counttable[1] + 1;
                //                                    }
                //                                }
                //                            }
                //                            nametable[name] = counttable;
                //                        }
                //                        break;
                //                    case "L7":
                //                        for (int k = 0; k < teacherlist.Length; k++)
                //                        {
                //                            int[] counttable = new int[11];
                //                            string name = "";
                //                            foreach (DictionaryEntry entry in nametable)
                //                            {
                //                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                //                                {
                //                                    counttable = (int[])entry.Value;
                //                                    name = entry.Key.ToString();
                //                                    if (k == 0)
                //                                    {
                //                                        counttable[2] = counttable[2] + 1;
                //                                    }
                //                                    else
                //                                    {
                //                                        counttable[3] = counttable[3] + 1;
                //                                    }
                //                                }
                //                            }
                //                            nametable[name] = counttable;
                //                        }
                //                        break;
                //                    case "M1":
                //                    case "M2":
                //                    case "M3":
                //                        for (int k = 0; k < teacherlist.Length; k++)
                //                        {
                //                            int[] counttable = new int[11];
                //                            string name = "";
                //                            foreach (DictionaryEntry entry in nametable)
                //                            {
                //                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                //                                {
                //                                    counttable = (int[])entry.Value;
                //                                    name = entry.Key.ToString();
                //                                    if (k == 0)
                //                                    {
                //                                        counttable[4] = counttable[4] + 1;
                //                                    }
                //                                    else
                //                                    {
                //                                        counttable[4] = counttable[4] + 1;
                //                                    }
                //                                }
                //                            }
                //                            nametable[name] = counttable;
                //                        }
                //                        break;
                //                    case "A1":
                //                    case "A2":
                //                        for (int k = 0; k < teacherlist.Length; k++)
                //                        {
                //                            int[] counttable = new int[11];
                //                            string name = "";
                //                            foreach (DictionaryEntry entry in nametable)
                //                            {
                //                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                //                                {
                //                                    counttable = (int[])entry.Value;
                //                                    name = entry.Key.ToString();
                //                                    if (k == 0)
                //                                    {
                //                                        counttable[7] = counttable[7] + 1;
                //                                    }
                //                                    else
                //                                    {
                //                                        counttable[7] = counttable[7] + 1;
                //                                    }
                //                                }
                //                            }
                //                            nametable[name] = counttable;
                //                        }
                //                        break;
                //                    case "A3":
                //                        for (int k = 0; k < teacherlist.Length; k++)
                //                        {
                //                            int[] counttable = new int[11];
                //                            string name = "";
                //                            foreach (DictionaryEntry entry in nametable)
                //                            {
                //                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                //                                {
                //                                    counttable = (int[])entry.Value;
                //                                    name = entry.Key.ToString();
                //                                    if (k == 0)
                //                                    {
                //                                        counttable[5] = counttable[5] + 1;
                //                                    }
                //                                    else
                //                                    {
                //                                        counttable[6] = counttable[6] + 1;
                //                                    }
                //                                }
                //                            }
                //                            nametable[name] = counttable;
                //                        }
                //                        break;
                //                    case "G1":
                //                        for (int k = 0; k < teacherlist.Length; k++)
                //                        {
                //                            int[] counttable = new int[11];
                //                            string name = "";
                //                            foreach (DictionaryEntry entry in nametable)
                //                            {
                //                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                //                                {
                //                                    counttable = (int[])entry.Value;
                //                                    name = entry.Key.ToString();
                //                                    if (k == 0)
                //                                    {
                //                                        counttable[8] = counttable[8] + 1;
                //                                    }
                //                                    else
                //                                    {
                //                                        counttable[8] = counttable[8] + 1;
                //                                    }
                //                                }
                //                            }
                //                            nametable[name] = counttable;
                //                        }
                //                        break;
                //                    case "G2":
                //                        for (int k = 0; k < teacherlist.Length; k++)
                //                        {
                //                            int[] counttable = new int[11];
                //                            string name = "";
                //                            foreach (DictionaryEntry entry in nametable)
                //                            {
                //                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                //                                {
                //                                    counttable = (int[])entry.Value;
                //                                    name = entry.Key.ToString();
                //                                    if (k == 0)
                //                                    {
                //                                        counttable[9] = counttable[9] + 1;
                //                                    }
                //                                    else
                //                                    {
                //                                        counttable[9] = counttable[9] + 1;
                //                                    }
                //                                }
                //                            }
                //                            nametable[name] = counttable;
                //                        }
                //                        break;
                //                    case "HL":
                //                        for (int k = 0; k < teacherlist.Length; k++)
                //                        {
                //                            int[] counttable = new int[11];
                //                            string name = "";
                //                            foreach (DictionaryEntry entry in nametable)
                //                            {
                //                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                //                                {
                //                                    counttable = (int[])entry.Value;
                //                                    name = entry.Key.ToString();
                //                                    if (k == 0)
                //                                    {
                //                                        counttable[10] = counttable[10] + 1;
                //                                    }
                //                                    else
                //                                    {
                //                                        counttable[10] = counttable[10] + 1;
                //                                    }
                //                                }
                //                            }
                //                            nametable[name] = counttable;
                //                        }
                //                        break;
                //                    default:
                //                        break;

                //                }
                //            }
                //        }
                //    }
                //}

                #endregion

                #region

                for (int i = 1; i < 38; i++)
                {
                    for (int j = 3; j < 22; j++)
                    {
                        if (worksheetschedule.Cells[j, i].Value.TextValue == null)
                        {
                            continue;
                        }
                        string value = worksheetschedule.Cells[j, i].Value.TextValue.Trim();
                        if ((value.Contains("L") || value.Contains("M") || value.Contains("A") || value.Contains("H") ||
                             value.Contains("G")) && value.Contains("-"))
                        {
                            //int start = value.IndexOf("(") > -1 ? value.IndexOf("(") : value.IndexOf("（");
                            //int end = value.IndexOf(")") > -1 ? value.IndexOf(")") : value.IndexOf("）");
                            var maches = System.Text.RegularExpressions.Regex.Match(value, ".+\\d");
                            string teacher = value.Substring(maches.Length).Trim();
                            string[] teacherlist = new string[] {};
                            if (teacher.Contains("+"))
                            {
                                teacherlist = teacher.Split('+');
                            }
                            else if (teacher.Contains("/"))
                            {
                                teacherlist = teacher.Split('/');
                            }
                            else
                            {
                                teacherlist = new string[1];
                                teacherlist[0] = teacher.Trim();
                            }

                            if (teacherlist.Length < 1)
                            {
                                continue;
                            }
                            else
                            {
                                string symbol = value.Substring(0, 2);
                                switch (symbol)
                                {
                                    case "L1":
                                    case "L2":
                                    case "L3":
                                    case "L4":
                                    case "L5":
                                    case "L6":
                                        for (int k = 0; k < teacherlist.Length; k++)
                                        {
                                            int[] counttable = new int[11];
                                            string name = "";
                                            foreach (DictionaryEntry entry in nametable)
                                            {
                                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                                                {
                                                    counttable = (int[]) entry.Value;
                                                    name = entry.Key.ToString();
                                                    if (k == 0)
                                                    {
                                                        counttable[0] = counttable[0] + 1;
                                                    }
                                                    else
                                                    {
                                                        counttable[1] = counttable[1] + 1;
                                                    }
                                                }
                                            }
                                            nametable[name] = counttable;
                                        }
                                        break;
                                    case "L7":
                                        for (int k = 0; k < teacherlist.Length; k++)
                                        {
                                            int[] counttable = new int[11];
                                            string name = "";
                                            foreach (DictionaryEntry entry in nametable)
                                            {
                                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                                                {
                                                    counttable = (int[]) entry.Value;
                                                    name = entry.Key.ToString();
                                                    if (k == 0)
                                                    {
                                                        counttable[2] = counttable[2] + 1;
                                                    }
                                                    else
                                                    {
                                                        counttable[3] = counttable[3] + 1;
                                                    }
                                                }
                                            }
                                            nametable[name] = counttable;
                                        }
                                        break;
                                    case "M1":
                                    case "M2":
                                    case "M3":
                                        for (int k = 0; k < teacherlist.Length; k++)
                                        {
                                            int[] counttable = new int[11];
                                            string name = "";
                                            foreach (DictionaryEntry entry in nametable)
                                            {
                                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                                                {
                                                    counttable = (int[]) entry.Value;
                                                    name = entry.Key.ToString();
                                                    if (k == 0)
                                                    {
                                                        counttable[4] = counttable[4] + 1;
                                                    }
                                                    else
                                                    {
                                                        counttable[4] = counttable[4] + 1;
                                                    }
                                                }
                                            }
                                            nametable[name] = counttable;
                                        }
                                        break;
                                    case "A1":
                                    case "A2":
                                        for (int k = 0; k < teacherlist.Length; k++)
                                        {
                                            int[] counttable = new int[11];
                                            string name = "";
                                            foreach (DictionaryEntry entry in nametable)
                                            {
                                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                                                {
                                                    counttable = (int[]) entry.Value;
                                                    name = entry.Key.ToString();
                                                    if (k == 0)
                                                    {
                                                        counttable[7] = counttable[7] + 1;
                                                    }
                                                    else
                                                    {
                                                        counttable[7] = counttable[7] + 1;
                                                    }
                                                }
                                            }
                                            nametable[name] = counttable;
                                        }
                                        break;
                                    case "A3":
                                        for (int k = 0; k < teacherlist.Length; k++)
                                        {
                                            int[] counttable = new int[11];
                                            string name = "";
                                            foreach (DictionaryEntry entry in nametable)
                                            {
                                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                                                {
                                                    counttable = (int[]) entry.Value;
                                                    name = entry.Key.ToString();
                                                    if (k == 0)
                                                    {
                                                        counttable[5] = counttable[5] + 1;
                                                    }
                                                    else
                                                    {
                                                        counttable[6] = counttable[6] + 1;
                                                    }
                                                }
                                            }
                                            nametable[name] = counttable;
                                        }
                                        break;
                                    case "G1":
                                    case "Gk":
                                        for (int k = 0; k < teacherlist.Length; k++)
                                        {
                                            int[] counttable = new int[11];
                                            string name = "";
                                            foreach (DictionaryEntry entry in nametable)
                                            {
                                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                                                {
                                                    counttable = (int[]) entry.Value;
                                                    name = entry.Key.ToString();
                                                    if (k == 0)
                                                    {
                                                        counttable[8] = counttable[8] + 1;
                                                    }
                                                    else
                                                    {
                                                        counttable[8] = counttable[8] + 1;
                                                    }
                                                }
                                            }
                                            nametable[name] = counttable;
                                        }
                                        break;
                                    case "G2":
                                        for (int k = 0; k < teacherlist.Length; k++)
                                        {
                                            int[] counttable = new int[11];
                                            string name = "";
                                            foreach (DictionaryEntry entry in nametable)
                                            {
                                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                                                {
                                                    counttable = (int[]) entry.Value;
                                                    name = entry.Key.ToString();
                                                    if (k == 0)
                                                    {
                                                        counttable[9] = counttable[9] + 1;
                                                    }
                                                    else
                                                    {
                                                        counttable[9] = counttable[9] + 1;
                                                    }
                                                }
                                            }
                                            nametable[name] = counttable;
                                        }
                                        break;
                                    case "HL":
                                        for (int k = 0; k < teacherlist.Length; k++)
                                        {
                                            int[] counttable = new int[11];
                                            string name = "";
                                            foreach (DictionaryEntry entry in nametable)
                                            {
                                                if (entry.Key.ToString().ToUpper() == teacherlist[k].Trim().ToUpper())
                                                {
                                                    counttable = (int[]) entry.Value;
                                                    name = entry.Key.ToString();
                                                    if (k == 0)
                                                    {
                                                        counttable[10] = counttable[10] + 1;
                                                    }
                                                    else
                                                    {
                                                        counttable[10] = counttable[10] + 1;
                                                    }
                                                }
                                            }
                                            nametable[name] = counttable;
                                        }
                                        break;
                                    default:
                                        break;

                                }
                            }
                        }
                    }
                }

                #endregion

                xtc.SelectedTabPageIndex = 1;

                worksheetcount = spreadsheetControl2.ActiveWorksheet;
                int row = 3;
                int col = 1;
                bool flag = true;
                while (flag)
                {
                    string over = worksheetcount.Cells[row, 0].Value.TextValue;
                    if (over.ToLower().Contains("total"))
                    {
                        flag = false;
                        break;
                    }
                    else
                    {
                        string fullname = worksheetcount.Cells[row, 0].Value.TextValue.Trim().ToLower();
                        foreach (DictionaryEntry entryname in NameListTable)
                        {
                            if (entryname.Value.ToString().Trim().ToLower().Contains(fullname))
                            {
                                string shortname = entryname.Key.ToString().Trim().ToLower();
                                foreach (DictionaryEntry dictionaryEntry in nametable)
                                {
                                    if (dictionaryEntry.Key.ToString().Trim().ToLower().Equals(shortname))
                                    {
                                        int[] tableInts = (int[]) dictionaryEntry.Value;
                                        for (int i = 0; i < 11; i++)
                                        {
                                            if (tableInts[i] == 0)
                                            {

                                            }
                                            else
                                            {
                                                worksheetcount.Cells[row, 1 + i].Value = tableInts[i];
                                            }

                                        }
                                    }
                                }
                            }
                        }
                        //string indexname = fullname.Trim().Substring(0, 2).ToLower();
                        //foreach (DictionaryEntry entry in nametable)
                        //{
                        //    if (entry.Key.ToString().Trim().Substring(0,2).ToLower().Equals(indexname))
                        //    {
                        //        int[] tableInts = (int[]) entry.Value;
                        //        for (int i = 0; i < 11; i++)
                        //        {
                        //            worksheetcount.Cells[row, 1 + i].Value = tableInts[i];
                        //        }
                        //    }
                        //}
                    }
                    row++;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btncheck_Click(object sender, EventArgs e)
        {
            Worksheet worksheet = spreadsheetControl1.ActiveWorksheet;
            //for (int i = 0; i < 22; i++)
            //{
            //    for (int j = 0; j < 38; j++)
            //    {
            //        string content = worksheet.Cells[i, j].Value.TextValue;
            //        if (!String.IsNullOrEmpty(content) && (content.Contains("(") || content.Contains("（")) && (content.IndexOf("L") == 0 || content.IndexOf("M") == 0 || content.IndexOf("A") == 0 || content.IndexOf("G") == 0 || content.IndexOf("H") == 0))
            //        {
            //            if (content.IndexOf("-") != 2)
            //            {
            //                worksheet.Cells[i, j].FillColor = DevExpress.Utils.DXColor.Red;
            //            }
            //        }
            //    }
            //}
            for (int i = 0; i < 22; i++)
            {
                for (int j = 0; j < 38; j++)
                {
                    string content = worksheet.Cells[i, j].Value.TextValue;
                    if (!String.IsNullOrEmpty(content)  && (content.IndexOf("L") == 0 || content.IndexOf("M") == 0 || content.IndexOf("A") == 0 || content.IndexOf("G") == 0 || content.IndexOf("H") == 0))
                    {
                        if (content.IndexOf("-") != 2)
                        {
                            worksheet.Cells[i, j].FillColor = DevExpress.Utils.DXColor.Red;
                        }
                        //if (content.Contains("Gk"))
                        //{
                        //    worksheet.Cells[i, j].FillColor = DevExpress.Utils.DXColor.Red;
                        //}
                    }
                }
            }
        }

        private void btnimportstas_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Excel文件";
            ofd.FileName = "";
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            ofd.RestoreDirectory = true;
            ofd.Filter = "Excel文件(*.xls)|*.xls|Excel2007文件(*.xlsx)|*.xlsx";
            ofd.ValidateNames = true; //文件有效性验证ValidateNames，验证用户输入是否是一个有效的Windows文件名
            ofd.CheckFileExists = true; //验证路径有效性
            ofd.CheckPathExists = true; //验证文件有效性
            string filePath = string.Empty;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;
            }

            if (!string.IsNullOrEmpty(filePath))
            {
                IWorkbook workbook = spreadsheetControl2.Document;
                workbook.LoadDocument(filePath);
            }

            xtc.SelectedTabPageIndex = 1;
        }

        private void btnexport_Click(object sender, EventArgs e)
        {
            string dir = System.Environment.CurrentDirectory;
            //spreadsheetControl2.SaveDocument("ww.xls");
            string localFilePath = "", fileNameExt = "", newFileName = "", FilePath = "";
            SaveFileDialog saveFileDialog = new SaveFileDialog();


            //设置文件类型
            //书写规则例如：txt files(*.txt)|*.txt
            saveFileDialog.Filter = "xls files(*.xls)|*.xls|xls 2007 files(*.xlsx)|*.xlsx|All files(*.*)|*.*";
            //设置默认文件名（可以不设置）
            saveFileDialog.FileName = "siling-Data";
            //主设置默认文件extension（可以不设置）
            saveFileDialog.DefaultExt = "xml";
            //获取或设置一个值，该值指示如果用户省略扩展名，文件对话框是否自动在文件名中添加扩展名。（可以不设置）
            saveFileDialog.AddExtension = true;

            //设置默认文件类型显示顺序（可以不设置）
            saveFileDialog.FilterIndex = 1;

            //保存对话框是否记忆上次打开的目录
            saveFileDialog.RestoreDirectory = true;

            // Show save file dialog box
            DialogResult result = saveFileDialog.ShowDialog();
            //点了保存按钮进入
            if (result == DialogResult.OK)
            {
                //获得文件路径
                localFilePath = saveFileDialog.FileName.ToString();
                spreadsheetControl2.SaveDocument(localFilePath);
            }
            MessageBox.Show("导出完成！");
        }
    }
}
