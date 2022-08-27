using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudyStatistic
{

    class Empolyee
    {
        //员工编号
        string id;
        //姓名
        string name;
        //部门
        string development;
        //职位
        string job;
        //当前状态
        string available;
        //学习的专题
        string zhuanti;
        //有两种 学习中，已完成
        string finished;
        string StartTime;
        string EndTime;
        string LastStudy;
        string studyTime;
        string mianshouTime;
        string TotalStudyTime;

        public Empolyee(string id, string name)
        {
            this.id = id;
            this.name = name;
        }
        public Empolyee(string id, string name, string development, string job, string available, string zhuanti, string finished, string startTime, string endTime, string lastStudy, string studyTime, string mianshouTime, string totalStudyTime)
        {
            this.id = id;
            this.name = name;
            this.development = development;
            this.job = job;
            this.available = available;
            this.zhuanti = zhuanti;
            this.finished = finished;
            StartTime = startTime;
            EndTime = endTime;
            LastStudy = lastStudy;
            this.studyTime = studyTime;
            this.mianshouTime = mianshouTime;
            TotalStudyTime = totalStudyTime;
        }

        public string Name { get => name; set => name = value; }
        public string Development { get => development; set => development = value; }
        public string Job { get => job; set => job = value; }
        public string Available { get => available; set => available = value; }
        public string Zhuanti { get => zhuanti; set => zhuanti = value; }
        public string Finished { get => finished; set => finished = value; }
        public string StartTime1 { get => StartTime; set => StartTime = value; }
        public string EndTime1 { get => EndTime; set => EndTime = value; }
        public string LastStudy1 { get => LastStudy; set => LastStudy = value; }
        public string StudyTime { get => studyTime; set => studyTime = value; }
        public string MianshouTime { get => mianshouTime; set => mianshouTime = value; }
        public string TotalStudyTime1 { get => TotalStudyTime; set => TotalStudyTime = value; }
        public string Id { get => id; set => id = value; }
    }
}
