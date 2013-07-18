using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleQueue
{
    class Foo
    {
        public static int Sum(int num)
        {
            int total = 0;
            for (int i = 1; i <= num; i++)
            {
                total += i;
            }
            return total;
        }
    }
}
