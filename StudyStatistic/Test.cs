using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using _Excel = Microsoft.Office.Interop.Excel;
namespace StudyStatistic
{

    class Test
    {
        //把名单和课程学习名单核对，筛选出没有学习的人，添加到Sheet1
        public static void NoStudy(_Excel.Worksheet sheet1, _Excel.Worksheet sheet2, ArrayList nostudy)
        {
            int rownum1 = sheet1.UsedRange.Rows.Count - 1;
            int rownum2 = sheet2.UsedRange.Rows.Count - 1;
        

            for (int i = 1; i < rownum1; i++)
            {
                string t1 = sheet1.Cells[3][i + 1].Value2;
                string name = sheet1.Cells[4][i + 1].Value2;
                string cleaned = Regex.Replace(t1, "[^0-9]", "");
                int t1v = 0;
                int.TryParse(cleaned, out t1v);
                int index = 1;
                for (int j = index; j < rownum2; j++)
                {
                    string t2 = sheet2.Cells[2][j+1].Value2;
                    t2 = t2.Replace("\t", String.Empty);
                    int t2v = 0;
                    int.TryParse(t2, out t2v);
                    //如果学习名单中有他的员工编号，那么就跳出内层循环，从下一个人开始找
                    if (t1v==t2v)
                    {
                        break;
                    }

                    //没找到他，把他加入名单
                    if (t1v>t2v)
                    {
                        nostudy.Add(new Empolyee(t1, name));
                    }
                }

            }
        }
        public static void NoStudy1(Dictionary<string, string> mingce, Dictionary<string, string> study, ArrayList nostudy) 
        {
            Dictionary<string, string>.KeyCollection keys1 = mingce.Keys;
            //对名册上的每一个id，去学习单中判断有没有他的id，没有就加入未学习名单
            foreach(string id1 in keys1) 
            {
                if (!study.ContainsKey(id1)) 
                {
                    nostudy.Add(new Empolyee(id1, mingce[id1]));
                }
            }
        }
        static void Main(string[] args)
        {
            //string s1 = "E0003901143";
            //string cleaned = Regex.Replace(s1, "[^0-9]", "");
            //int value = 0;
            //int.TryParse(cleaned, out value);

            //Console.WriteLine(value);
            Dictionary<string, string> excel1IdAndName = new Dictionary<string, string>();
            Dictionary<string, string> excel2IdAndName = new Dictionary<string, string>();
            Dictionary<string, string> nostudyIdAndName = new Dictionary<string, string>();
            Excel excel1 = new Excel("C:\\Users\\Administrator\\Desktop\\学习情况统计\\简洁花名册副本.xlsx", 1);
            Excel excel2 = new Excel("C:\\Users\\Administrator\\Desktop\\学习情况统计\\学习副本.csv", 1);
            
            ArrayList nostudy = new ArrayList();
            excel1.ReadIdAndName(excel1IdAndName);
            excel2.ReadIdAndName(excel2IdAndName);
            NoStudy1(excel1IdAndName, excel2IdAndName, nostudy);
            //excel1.Sort("C");
            //excel2.Sort("B");
            //NoStudy(excel1.GetWorksheet(), excel2.GetWorksheet(), nostudy);
            Excel res = new Excel("C:\\Users\\Administrator\\Desktop\\学习情况统计\\Result1.xlsx");
            res.fillworksheet(nostudy);
            excel1.exit();
            excel2.exit();
            res.exit();
        }
        
    }
}
