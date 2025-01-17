﻿using System;
using System.IO;
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
        public static void tongji2(Dictionary<string,Empolyee> huamingce,Dictionary<string,Empolyee> xxjl,ArrayList department,Dictionary<string,int> inprocess,Dictionary<string,int> finished,Dictionary<string,int> nostart) 
        {
            foreach(Empolyee ep in huamingce.Values) 
            {
                //如果学习记录当中没有他的员工编码
                if (!xxjl.ContainsKey(ep.Id)) 
                {
                    //如果未开始Dictionary中没有这个部门，那么新建这个键
                    if (!nostart.ContainsKey(ep.Development)) 
                    {
                        nostart.Add(ep.Development, 1);
                    }
                    else
                    {
                        nostart[ep.Development] += 1;
                    }
                }
                //如果学习记录当中有他的员工编码
                else
                {
                    //判断他在学习记录当中的状态
                    if (xxjl[ep.Id].Finished.Equals("学习中")) 
                    {
                        //如果未完成字典中不包含这个部门
                        if (!inprocess.ContainsKey(ep.Development)) 
                        {
                            inprocess.Add(ep.Development, 1);
                        }
                        else
                        {
                            inprocess[ep.Development] += 1;
                        }
                    }
                    //代表他的学习记录是已完成了
                    else
                    {
                        if (!finished.ContainsKey(ep.Development)) 
                        {
                            finished.Add(ep.Development, 1);
                        }
                        else
                        {
                            finished[ep.Development] += 1;
                        }
                    }
                }
            }
        }
        //1.首先，部门按照花名册的来算
        //2.其次，不需要统计学习名单当中专题为"-"的员工
        //3.在花名册当中，但是却没有出现在学习名单中的员工，算作没有开始学习
        //所以，第一步是把花名册当中的所有员工读入一个Dictionary，key是员工编码，value是Employee，包含姓名和部门
        //第一步当中还会生成一个部门的ArrayList，保存所有的部门
        //第二步，把学习名单中，除了专题为"-"的，其他的员工都读入Dictionary，key是员工编码，value是Employee，包含所有属性
        //第三步，按照部门统计，对于ArrayList每一个部门，用花名册的Dictionary当中的Key
        //去学习名单当中的Dictionary去找他的学习状态，如果找不到，算作这个部门未完成的+1
        //如果找到了，则根据学习状态来判断是未完成的+1，还是完成的+1
        static void Main(string[] args)
        {
            string zhuanti = null;
            DateTime dateTime = DateTime.UtcNow.Date;
            dateTime.AddHours(8.0);
            string date = dateTime.ToString("yyyy-MM-dd");
            Console.WriteLine("当前exe所在目录为" + Directory.GetCurrentDirectory());
            string[] files = Directory.GetFiles(Directory.GetCurrentDirectory());
            ArrayList realfile = new ArrayList();
            int rownum = 1;
            foreach(string file in files) 
            {
                if (file.Contains("xlsx") || file.Contains("csv")) 
                {
                    Console.WriteLine(rownum+" "+file);
                    realfile.Add(file);
                    rownum++;
                }
            }
            string huamingcePath;
            string xuexijiluPath;
            string StatisticPath;
            string mingxiPath;
            int select;
            Console.WriteLine("下面要你输入文件的路径，输入数字即可，按回车继续");
            Console.WriteLine("请选择花名册的路径");
            select = Convert.ToInt32(Console.ReadLine());
            huamingcePath = realfile[select - 1].ToString();
            Console.WriteLine("请选择学习明细路径");
            select = Convert.ToInt32(Console.ReadLine());
            xuexijiluPath = realfile[select - 1].ToString();


            Console.WriteLine("请静等1分钟左右");
            Dictionary<string, Empolyee> hmc = new Dictionary<string, Empolyee>();
            Dictionary<string, Empolyee> xxjl = new Dictionary<string, Empolyee>();
            Dictionary<string, int> inprocess = new Dictionary<string, int>();
            Dictionary<string, int> finished = new Dictionary<string, int>();
            Dictionary<string, int> notstart = new Dictionary<string,int>();
            ArrayList Department = new ArrayList();
            //1.打开花名册
            Excel huamingce = new Excel(huamingcePath,1);
            //2.读取到Dictionary和部门的ArrayList
            huamingce.ReadeHuamingce(hmc, Department);
            //3.打开学习记录
            Excel record = new Excel(xuexijiluPath, 1);
            //4.把学习记录读入到Dictionary,并且看一下当前项目是什么专题的
            record.ReadEmolyee(xxjl,ref zhuanti);
            //5.开始统计各个部门的完成率
            tongji2(hmc, xxjl, Department, inprocess, finished,notstart);
            //6.准备输出结果
            StatisticPath = Directory.GetCurrentDirectory() + "\\" + zhuanti + "学习统计.xlsx";
            mingxiPath = Directory.GetCurrentDirectory() + "\\" + zhuanti + "学习统计明细_" + date + ".xlsx";
            Excel res = new Excel(StatisticPath);
            Excel res2 = new Excel(mingxiPath);
            res.FillStatistic(finished,inprocess,notstart,Department);
            res2.FillMingxi(hmc, xxjl);
            huamingce.exit();
            record.exit();
            res.exit();
            res2.exit();
            Console.WriteLine("统计结果和明细已经都打印到了指定位置了");
        }
        
    }
}
