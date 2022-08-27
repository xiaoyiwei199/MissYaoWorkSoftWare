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
        //找到的是没有开始学习的人
        public static void NoStudy1(Dictionary<string, string> mingce, Dictionary<string, Empolyee> study, ArrayList nostudy) 
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
        //这些都是有学习数据的人，但是有的人是学习中，有的人是已完成，统计他们的数据即可
        //统计每个部门名称，应学人数，完成人数，完成率，未完成人数，
        public static void tongji(Dictionary<string, Empolyee> study) 
        {
            
        }
        static void Main(string[] args)
        {
            //string s1 = "E0003901143";
            //string cleaned = Regex.Replace(s1, "[^0-9]", "");
            //int value = 0;
            //int.TryParse(cleaned, out value);

            //Console.WriteLine(value);

            Dictionary<string, string> excel1IdAndName = new Dictionary<string, string>();
            Dictionary<string, Empolyee> excel2IdAndName = new Dictionary<string, Empolyee>();
            Dictionary<string, Empolyee> nostudyIdAndName = new Dictionary<string, Empolyee>();
            Dictionary<string, Empolyee> haveStudy = new Dictionary<string, Empolyee>();
            Excel excel1 = new Excel("C:\\Users\\Administrator\\Desktop\\学习情况统计\\简洁花名册副本.xlsx", 1);
            //学习副本里全覆盖Employee这些属性，但是花名册当中并没有那些属性。
            //我目前要做的就是统计，excel1只是用来筛选出没有开始学习的人的
            //剩下的工作都是excel2当中做的，复制过去即可
            Excel excel2 = new Excel("C:\\Users\\Administrator\\Desktop\\学习情况统计\\结果副本.xlsx", 1);
            
            ArrayList nostudy = new ArrayList();
            ArrayList Department = new ArrayList();
            excel1.ReadIdAndName(excel1IdAndName);
            excel2.ReadEmolyee(excel2IdAndName,Department);
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
