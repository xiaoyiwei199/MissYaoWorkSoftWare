using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudyStatistic
{
    class Empolyee
    {
        string id;
        string name;

        public Empolyee(string id,string name) 
        {
            this.id = id;
            this.name = name;
        }

        public string Id { get => id; set => id = value; }
        public string Name { get => name; set => name = value; }
    }
}
