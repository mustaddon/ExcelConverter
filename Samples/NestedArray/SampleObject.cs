using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NestedArray
{
    class SampleObject
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public SampleUser User { get; set; }
        public SampleChild[] Childs { get; set; }
    }

    class SampleUser
    {
        public string Login { get; set; }
        public string Name { get; set; }
    }

    public class SampleChild
    {
        public string Prop1 { get; set; }
        public string Prop2 { get; set; }
        public string[] Prop3 { get; set; }
    }
}
