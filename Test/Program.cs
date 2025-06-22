using System.Diagnostics;
using System.Net;
using System.Runtime.InteropServices;
using NewLife;
using NewLife.Office;
using NewLife.Log;
using NewLife.Reflection;

namespace Test;

class Program
{
    static void Main(String[] args)
    {
        Runtime.CreateConfigOnMissing = false;
        XTrace.UseConsole();

        try
        {
            Test1();
        }
        catch (Exception ex)
        {
            XTrace.WriteException(ex);
        }

        Console.WriteLine("OK!");
        Console.ReadKey();
    }

    static void Test1()
    {
    }
}