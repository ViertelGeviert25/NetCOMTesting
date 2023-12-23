using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace BinaryTests
{
    [Serializable]
    class BinaryExample
    {
        public int Height { get; set; }
        public int Width { get; set; }
        public byte[] Values { get; set; }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            var example = new BinaryExample();
            example.Height = 25;
            example.Width = 30;
            example.Values = new byte[10] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };

            //using System.Runtime.Serialization.Formatters.Binary; - mscorlib.dll
            var fileName = @"H:\testdata\example.bin"; //"example.bin";
            using (var stream = File.Open(fileName, FileMode.Create))
            {
                var binaryFormatter = new BinaryFormatter();
                binaryFormatter.Serialize(stream, example);
            }
        }
    }
}
