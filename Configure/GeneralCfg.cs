namespace DingExcelTool.Configure
{
    using System;
    using System.Collections.Generic;
    using DingExcelTool.Data;

    internal static class GeneralCfg
    {
        public const string DefaultDataPath = "Data/DefaultData.Json";
        public const string CustomDataPath = "Data/CustomData.Json";
        public const string ProtocPath = "tools/protoc.exe";

        public const string ProtoMetaFileSuffix = "PM.pbmeta";
        public const string ProtoDataFileSuffix = "PD.pbdata";
        public static string ExcelScriptFileSuffix(ScriptTypeEn scriptType)
        {
            string fileSuffix = "Excel";
            return scriptType switch
            {
                ScriptTypeEn.CSharp => $"{fileSuffix}.cs",
                _ => throw new NotImplementedException($"未知的脚本类型：{scriptType}")
            };
        }

        public readonly static string[] ExcelBaseType = { "int", "long", "double", "bool", "string" };
        public const string LocalizationTxtSymbol = "%string";
        public const string LocalizationImgSymbol = "%%string";
        public const string IndependentKeySymbol = "*";
        public const string UnionKeySymbol = "**";

        public static readonly Dictionary<string, string> BaseType2ProtoMap = new()
        { 
            {"int", "int32"},
            {"long", "int64"},
            {"double", "double"},
            {"bool", "bool"},
            {"string", "string"},
        };

        public const string ProtoMetaPackageName = "Business.Data.Excel";

        public const string ExcelScriptTemplateFileName = "{0}_{1}_ExcelSccriptTemplate.txt";
    }
}
