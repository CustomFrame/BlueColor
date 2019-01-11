eg:
```C#
using System;

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
```
Console
```text 
你好,世界!
```
HelloLogFile.Info.log
```text
2019-01-11 21:05:34.0027|INFO|HelloLogFile|你好,世界!
```
