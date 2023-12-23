using MyCOMLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TDAPIOLELib;

namespace ConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var mc = new MyClass();
            mc.MyMethod("test");
            Console.ReadKey();
        }
    }
}
