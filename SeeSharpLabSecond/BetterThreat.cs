using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeeSharpLabSecond
{
    class BetterThreat
    {
        public string Id { set; get; }
        public string Name { set; get; }
        public BetterThreat(string id, string name)
        {
            this.Id = id;
            this.Name = name;
        }
        public int GetId()
        {
            return int.Parse(Id.Substring(4));
        }
    }
}
