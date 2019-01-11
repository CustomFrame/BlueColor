﻿using System;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            var result = BlueColor.TestLog.HelloWorld.SayHello("世界");
            Console.WriteLine(result);

            var log = BlueColor.TestLog.BcLog.GetLogger("HelloLogFile");
            log.Info(result);
        }
    }
}
