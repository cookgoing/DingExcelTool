namespace DingExcelTool.Utils;

using System;
using System.Text;
using System.IO.Hashing;
using Data;
using ScriptHandler;

internal static class LanguageUtils
{
    public static int Hash32Bytes(string input)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));

        return (int)XxHash32.HashToUInt32(Encoding.UTF8.GetBytes(input));
    }
    
    public static IScriptLanguageHandler GetScriptLanguageHandler(ScriptTypeEn scriptType)
    {
        return scriptType switch
        {
            ScriptTypeEn.CSharp => CSharpLanguageHandler.Instance,
            _ => throw new Exception($"[GetScriptLanguageHandler]. 未知的脚本类型：{scriptType}")
        };
    }
}