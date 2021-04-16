using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LichThucTap
{
    public class OneWeek
    {
        public OneWeek()
        {
            T2 = new OneDay();
            T3 = new OneDay();
            T4 = new OneDay();
            T5 = new OneDay();
            T6 = new OneDay();
            T7 = new OneDay();
            CN = new OneDay();
        }
        public OneDay T2 { get; set; }
        public OneDay T3 { get; set; }
        public OneDay T4 { get; set; }
        public OneDay T5 { get; set; }
        public OneDay T6 { get; set; }
        public OneDay T7 { get; set; }
        public OneDay CN { get; set; }
    }
}
