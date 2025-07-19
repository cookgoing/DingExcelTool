using System.Collections.Generic;

namespace DingExcelTool.Configure
{
    using Data;
    
    internal static class LanguageCfg
    {
        public static readonly string LanguageExcelName = "Lanuage_{0}.xlsx";
        public static readonly string ImageExcelName = "Image_{0}.xlsx";
        public static readonly string LanguageTextImageReplaceArg = "hashID";
        public static readonly string LanguageExcelSheetName = "Sheet1";
        public static readonly string ImageExcelSheetName = "Sheet1";

        public static readonly string[] LanguageExcelNameArr = { "HashID", "Source", "Text"};
        public static readonly string[] ImageExcelNameArr = { "HashID", "Source", "Image"};
        public static readonly Dictionary<string, (string type, string platform, string comment)> LanguageExcelHeadDic = new()
        {
            {"HashID", ("*int", "c", "hash值")},
            {"Source", ("", "", "原始文本")},
            {"Text", ("string", "c", "翻译后的文本")},
        };
        public static readonly Dictionary<string, (string type, string platform, string comment)> ImageExcelHeadDic = new()
        {
            {"HashID", ("*int", "c", "hash值")},
            {"Source", ("", "", "原始路径")},
            {"Image", ("string", "c", "本地化后的新路径")},
        };
        
        public const string OuterLanguageHandlerInterfacePath = "Data/Language/OuterLanguageHandler.cs";
        public const string UILanguageHandlerPath = "Data/Language/UILanguageHandler.cs";
        public const string CodeLanguageHandlerPath = "Data/Language/CodeLanguageHandler.cs";
        public const string UILanguageHandlerClassName = "UILanguageHandler";
        public const string CodeLanguageHandlerClassName = "CodeLanguageHandler";
    }
}
