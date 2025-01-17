﻿using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using _Excel = Microsoft.Office.Interop.Excel;
namespace StudyStatistic
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel(string path,int Sheet) 
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }
        public _Excel.Workbook GetWorkbook() 
        {
            return wb;
        }
        //在路径上创造一个新的Excel文件
        public Excel(string path) 
        {
            wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            ws = wb.Worksheets[1];
            wb.SaveAs(path);
        }
        public _Excel.Worksheet GetWorksheet() 
        {
            return ws;
        }
        //对于csv来说，这种Sort会改变其数值
        public void Sort(string colCharacter) 
        {
            //用了几行
            int row = ws.UsedRange.Rows.Count;
            string srow = row.ToString();
            string range1 = colCharacter + "2";
            string range2 = colCharacter + srow;
            ws.Sort.SetRange(ws.Range[range1,range2]);
            ws.Sort.SortFields.Add(ws.Range[range1,range2], _Excel.XlSortOn.xlSortOnValues, _Excel.XlSortOrder.xlAscending,_Excel.XlSortMethod.xlPinYin);
            ws.Sort.Apply();
        }
        public string ReadCell(int row,int col) 
        {
            row++;
            col++;

            if (ws.Cells[row, col].Value2!= null) 
            {
                return ws.Cells[row, col].Value2;
            }
            else
            {
                return "";
            }
        }
        public void WriteCell(int row,int col,string words) 
        {
            row++;
            col++;

            ws.Cells[row, col].Value2 = words;
        }
        public string GetCell(int row,int col) 
        {
            row++;
            col++;
            return ws.Cells[row, col].Value2;
        }

        //读取更多的信息，封装到字典里，这次不只是名字
        //把这个excel的当前表的信息读到字典里，key是员工编号，值是这个员工 封装
        internal void ReadEmolyee(Dictionary<string, Empolyee> excel2IdAndName,ref string Zhuanti)
        {
            int colnum = ws.UsedRange.Columns.Count;
            int rownum = ws.UsedRange.Rows.Count;
            _Excel.Range allRange = ws.Range[ws.Cells[2,1],ws.Cells[rownum,colnum]];
            string name;
            string id;
            string department;
            string shortdepartment;
            string post;
            string condition;
            string zhuanti;
            string xuexicondition;
            string StartTime;
            string FinishTime;
            string LastStudyTime;
            string StudyTime;
            string FaceStudyTime;
            string TotalTime;
            string[] d;
            foreach ( _Excel.Range row in allRange.Rows) 
            {
                var resizerow = row.Resize[1, colnum];
                object[,] Arow = resizerow.Value2;
                zhuanti = Arow[1, 6].ToString().Replace("\t", String.Empty);
                //只有专题不为空的才有统计的必要
                if (!zhuanti.Equals("-"))
                {
                    Zhuanti = zhuanti;
                    name = Arow[1, 1].ToString().Replace("\t", String.Empty);
                    id = Arow[1, 2].ToString().Replace("\t", String.Empty);
                    if (id.Length != 11) 
                    {
                        id = "E00" + id;
                    }
                    department = Arow[1, 3].ToString().Replace("\t", String.Empty);
                    if (department.Contains("/"))
                    {
                        d = department.Split("/");
                        shortdepartment = d[1];
                    }
                    else
                    {
                        shortdepartment = department;
                    }
                    if (shortdepartment.Equals("法律事务部")) 
                    {
                        shortdepartment = "综合部（董事会办公室、法律事务部）";
                    }
                    post = Arow[1, 4].ToString().Replace("\t", String.Empty);
                    condition = Arow[1, 5].ToString().Replace("\t", String.Empty);
                    xuexicondition = Arow[1, 7].ToString().Replace("\t", String.Empty);
                    StartTime = Arow[1, 8].ToString().Replace("\t", String.Empty);
                    FinishTime = Arow[1, 9].ToString().Replace("\t", String.Empty);
                    LastStudyTime = Arow[1, 10].ToString().Replace("\t", String.Empty);
                    StudyTime = Arow[1, 11].ToString().Replace("\t", String.Empty);
                    FaceStudyTime = Arow[1, 12].ToString().Replace("\t", String.Empty);
                    TotalTime = Arow[1, 13].ToString().Replace("\t", String.Empty);
                    //如果是第一次加入这个Dictionary
                    if (!excel2IdAndName.ContainsKey(id)) 
                    { 
                    excel2IdAndName.Add(id, new Empolyee(id, name, shortdepartment, post, condition, zhuanti, xuexicondition, StartTime, FinishTime, LastStudyTime, StudyTime, FaceStudyTime, TotalTime));
                    }
                    //如果发现了他的第二条学习记录
                    else 
                    {
                        if (xuexicondition.Equals("已完成")) 
                        {
                            excel2IdAndName[id] = new Empolyee(id,name,shortdepartment,post,condition,zhuanti,xuexicondition,StartTime,FinishTime,LastStudyTime,StudyTime,FaceStudyTime);
                        }
                    }
                }
                
            }     
        }
        //统计学习明细：员工编码，姓名，公司邮箱，部门名称，学习状态，学习时长，开始学习时间，完成时间，上次学习时间
        //所有在花名册上的人都会被统计进去
        //如果花名册上的员工编号，在学习记录中找不到，那么他就是没开始学习，
        internal void FillMingxi(Dictionary<string, Empolyee> hmc, Dictionary<string, Empolyee> xxjl)
        {
            WriteCell(0, 0, "员工编码");
            WriteCell(0, 1, "姓名");
            WriteCell(0, 2, "公司邮箱");
            WriteCell(0, 3, "正式部门名称");
            WriteCell(0, 4, "学习时长");
            WriteCell(0, 5, "学习状态");
            WriteCell(0, 6, "开始学习时间");
            WriteCell(0, 7, "完成时间");
            WriteCell(0, 8, "上次学习时间");
            //2.建立写入WorkSheet的行和列的索引
            int row = 2;//从第row行开始写
            int colstart = 1;//从第colstart列开始写
            int colend = 9;//写到colend
            _Excel.Range range;
            foreach(Empolyee ep in hmc.Values) 
            {
                 range = ws.Range[ws.Cells[row, colstart], ws.Cells[row, colend]];
                var onerow = range.Resize[1, 9];
                //这位员工没有开始学习,写入到excel当中，一行一行的写
                if (!xxjl.ContainsKey(ep.Id)) 
                {
                    onerow[1, 1] = ep.Id;
                    onerow[1, 2] = ep.Name;
                    onerow[1, 3] = ep.Email;
                    onerow[1, 4] = ep.Development;
                    onerow[1, 5] = "未开始";
                    onerow[1, 6] = "-";
                    onerow[1, 7] = "-";
                    onerow[1, 8] = "-";
                    onerow[1, 9] = "-";
                }
                //这位员工在学习中或者是已经完成学习了
                else
                {
                    onerow[1, 1] = ep.Id;
                    onerow[1, 2] = ep.Name;
                    onerow[1, 3] = ep.Email;
                    onerow[1, 4] = ep.Development;
                    Empolyee tmp = xxjl[ep.Id];
                    onerow[1, 5] = tmp.Finished;
                    onerow[1, 6] = tmp.StudyTime;
                    onerow[1, 7] = tmp.StartTime1;
                    onerow[1, 8] = tmp.EndTime1;
                    onerow[1, 9] = tmp.LastStudy1;
                }
                row++;
            }

        }

        internal void ReadeHuamingce(Dictionary<string, Empolyee> hmc, ArrayList Department)
        {
            int colnum = ws.UsedRange.Columns.Count;
            int rownum = ws.UsedRange.Rows.Count;
            _Excel.Range allRange = ws.Range[ws.Cells[2, 1], ws.Cells[rownum, colnum]];
            string name;
            string id;
            string department;
            string shortdepartment;
            string[] d;
            string email;
            foreach (_Excel.Range row in allRange.Rows) 
            {
                var resizerow = row.Resize[1, colnum];
                object[,] Arow = resizerow.Value2;
                id = Arow[1, 3].ToString();
                name = Arow[1, 4].ToString();
                email = Arow[1, 5].ToString();
                department = Arow[1, 6].ToString();
                if (department.Contains("/")) 
                {
                    d = department.Split("/");
                    shortdepartment = d[1];
                }
                else
                {
                    shortdepartment = department;
                }
                if (shortdepartment.Equals("法律事务部")||shortdepartment.Equals("综合部（董事会办公室、法律事务部）"))
                {
                    shortdepartment = "综合部";
                }
                if (shortdepartment.Equals("中移物联网有限公司")) 
                {
                    shortdepartment = "公司领导";
                }
                if (shortdepartment.Equals("党委办公室（党群工作部、巡察工作办公室、党风廉政办公室）")) 
                {
                    shortdepartment = "党委办公室";
                }
                if (!Department.Contains(shortdepartment))
                {
                    Department.Add(shortdepartment);
                }

                hmc.Add(id, new Empolyee(id, name, shortdepartment,email));
            }
        }

        public void save() 
        {
            wb.Saved = true;
        }
        public void exit()
        {
            excel.DisplayAlerts = false;
            wb.Save();
            wb.Close(true);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);


        }
        public void addSheet(string name) 
        {
            _Excel.Worksheet newWorksheet;
            newWorksheet = wb.Worksheets.Add(After:ws);
            newWorksheet.Name = name;
        }

        public void fillworksheet(ArrayList res) 
        {
            //因为有标题头，都得从第二行第二列开始写
            int i = 1;
            int j = 1;
            foreach(Empolyee employee in res) 
            {
                WriteCell(i , j , employee.Id);
                WriteCell(i , j + 1, employee.Name);
                i++;
                
            }
        }
        public void FillStatistic(Dictionary<string,int> finished,Dictionary<string,int> inprocess,Dictionary<string,int> nostart,ArrayList department)  
        {
            //1.填写表头
            WriteCell(0, 0,"部门正式名称");
            WriteCell(0, 1, "未开始");
            WriteCell(0, 2, "学习中");
            WriteCell(0, 3, "已完成");
            WriteCell(0, 4, "应学人数");
            WriteCell(0, 5,"完成率");
            int xuexizhong;
            int wancheng;
            int total;
            int weikaishi;
            double percent;
            string p;
            int row = 1;
            int col = 0;
            foreach(string dp in department) 
            {
                //对于xuexizhong和wancheng，其实不一定有值的
                //可能这个部门都学完了，没有在学习中的
                //也有可能这个部门都在学习中，没有学完的
                xuexizhong = inprocess.ContainsKey(dp) ? inprocess[dp] : 0;
                wancheng = finished.ContainsKey(dp) ? finished[dp] : 0;
                weikaishi = nostart.ContainsKey(dp)? nostart[dp]:0;
                total = xuexizhong + wancheng+weikaishi;
                percent =(double) wancheng / total;
                p = String.Format("{0:P2}",percent);
                WriteCell(row, col, dp);
                WriteCell(row, col + 1, weikaishi.ToString());
                WriteCell(row, col + 2, xuexizhong.ToString());
                WriteCell(row, col + 3, wancheng.ToString());
                WriteCell(row, col + 4, total.ToString());
                WriteCell(row,col+5,p);
                row++;
            }
        }
    }
}
