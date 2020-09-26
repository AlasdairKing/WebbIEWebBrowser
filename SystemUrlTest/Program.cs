using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SystemUrlTest
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Uri uri = new Uri("https://www.clarosoftware.com/this/that/whatever.php?a=5&b=5#main");
            Console.WriteLine(uri.Host);
        }
    }
}
