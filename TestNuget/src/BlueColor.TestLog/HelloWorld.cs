using System;

namespace BlueColor.TestLog
{
    /// <summary>
    /// HelloWorld
    /// </summary>
    public class HelloWorld
    {
        /// <summary>
        /// 问候
        /// </summary>
        /// <param name="nickName">昵称</param>
        public static string SayHello(string nickName)
        {
            return $"Hello,{nickName}!";
        }
    }
}