using DingExcelTool.Utils;

namespace DingExcelTool.ScriptHandler
{
    using System;
    using System.IO;
    using System.Text;
    using System.Threading.Tasks;
    using System.Collections.Concurrent;
    using Data;
    using Configure;

    internal class CSharpSpecialExcelHandler : Singleton<CSharpSpecialExcelHandler>, IScriptSpecialExcelHandler
    {
        public async Task GenerateErrorCodeScript(ConcurrentDictionary<string, ErrorCodeScriptInfo> errorCodeHeadDic, string frameOutputFile, string businessOutputFile)
        {
            if (string.IsNullOrEmpty(frameOutputFile) || string.IsNullOrEmpty(businessOutputFile)) throw new Exception("[GenerateErrorCodeScript]. ErrorCode 需要有两个输出路径");
            if (errorCodeHeadDic == null) throw new Exception("[GenerateErrorCodeScript]. Errorcode 表，没有头部信息");

            string frameOutputDir = Path.GetDirectoryName(frameOutputFile);
            string businessOutputDir = Path.GetDirectoryName(businessOutputFile);
            if (!Directory.Exists(frameOutputDir)) Directory.CreateDirectory(frameOutputDir);
            if (!Directory.Exists(businessOutputDir)) Directory.CreateDirectory(businessOutputDir);

            StringBuilder frameFieldSB = new();
            StringBuilder businessFieldSB = new();
            foreach (ErrorCodeScriptInfo errorCodeInfo in errorCodeHeadDic.Values)
            {
                bool isFrame = errorCodeInfo.SheetName == SpecialExcelCfg.ErrorCodeFrameSheetName;
                foreach (ErrorCodeScriptFieldInfo fieldInfo in errorCodeInfo.Fields)
                {
                    if (isFrame) frameFieldSB.Append($"\t\tpublic const int {fieldInfo.CodeStr} = {fieldInfo.Code};").AppendLine(string.IsNullOrEmpty(fieldInfo.Comment) ? null : "//" + fieldInfo.Comment);
                    else businessFieldSB.Append($"\t\tpublic const int {fieldInfo.CodeStr} = {fieldInfo.Code};").AppendLine(string.IsNullOrEmpty(fieldInfo.Comment) ? null : "//" + fieldInfo.Comment);
                }
            }

            await WriteErrorCodeScript(frameOutputFile, frameFieldSB.ToString(), SpecialExcelCfg.ErrorCodeFramePackageName);
            await WriteErrorCodeScript(businessOutputFile, businessFieldSB.ToString(), SpecialExcelCfg.ErrorCodeBusinessPackageName);
        }

        public async Task GenerateSingleScript(SingleExcelHeadInfo singleHeadInfo, string outputFile, bool isClient)
        {
            if (string.IsNullOrEmpty(outputFile)) throw new Exception("[GenerateSingleScript]. 没有输出路径");

            string dirPath = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(dirPath)) Directory.CreateDirectory(dirPath);

            PlatformType platform = isClient ? PlatformType.Client : PlatformType.Server;
            StringBuilder fieldSB = new();
            foreach (SingleExcelFieldInfo fieldInfo in singleHeadInfo.Fields)
            {
                if ((platform & fieldInfo.Platform) == 0) continue;

                string filedValue = null;
                if (ExcelUtil.IsArrType(fieldInfo.Type))
                {
                    StringBuilder sb = new();
                    sb.Append('{');
                    string baseType = fieldInfo.Type.Substring(0, fieldInfo.Type.Length - 2);
                    string[] valueStrArr = fieldInfo.Value.Split(SpecialExcelCfg.SingleArrMapSplitSymbol);
                    foreach (string str in valueStrArr)
                    {
                        string value = CSharpExcelHandler.Instance.ExcelType2ScriptType(baseType, str).ToString();
                        if (baseType == "string") value = $"\"{value}\"";
                        else if (baseType == "bool") value = value?.ToLower();
                        sb.Append(value).Append(',');    
                    }
                    sb.Remove(sb.Length - 1, 1).Append('}');
                    filedValue  = sb.ToString();
                }
                else if (ExcelUtil.IsMapType(fieldInfo.Type))
                {
                    StringBuilder sb = new("new(){");
                    string innerTypes = fieldInfo.Type.Substring(4, fieldInfo.Type.Length - 5);
                    string[] keyValue = innerTypes.Split(',');
                    string kType = keyValue[0], vType = keyValue[1];
                    string[] valueStrArr = fieldInfo.Value.Split(SpecialExcelCfg.SingleArrMapSplitSymbol);
                    fieldInfo.Type = $"map<{kType},{vType}>";
                    foreach (string kvStr in valueStrArr)
                    {
                        string[] kvStrArr = kvStr.Split(SpecialExcelCfg.SingleMapKVSplitSymbol);
                        string kValue = CSharpExcelHandler.Instance.ExcelType2ScriptType(kType, kvStrArr[0]).ToString();
                        string vValue = CSharpExcelHandler.Instance.ExcelType2ScriptType(vType, kvStrArr[1]).ToString();
                        if (kType == "string") kValue = $"\"{kValue}\"";
                        else if (kType == "bool") kValue = kValue?.ToLower();
                        if (vType == "string") vValue = $"\"{vValue}\"";
                        else if (vType == "bool") vValue = vValue?.ToLower();
                        
                        sb.Append('{').Append(kValue).Append(',').Append(vValue).Append('}').Append(',');
                    }
                    sb.Remove(sb.Length - 1, 1).Append('}');
                    filedValue  = sb.ToString();
                }
                else
                {
                    filedValue = CSharpExcelHandler.Instance.ExcelType2ScriptType(fieldInfo.Type, fieldInfo.Value).ToString();
                    if (fieldInfo.Type == "string") filedValue = $"\"{filedValue}\"";
                    else if (fieldInfo.Type == "bool") filedValue = filedValue?.ToLower();
                }

                fieldSB.Append($"\t\tpublic readonly static {CSharpExcelHandler.Instance.ExcelType2ScriptTypeStr(fieldInfo.Type)} {fieldInfo.Name} = {filedValue};").AppendLine(string.IsNullOrEmpty(fieldInfo.Comment) ? null : "//" + fieldInfo.Comment);
            }

            await WriteSingleExcelScript(singleHeadInfo.ScriptName, outputFile, fieldSB.ToString());
        }

        private async Task WriteErrorCodeScript(string outputPath, string fieldStr, string packageName)
        {
            using StreamWriter sw = new StreamWriter(outputPath);
            StringBuilder scriptSB = new();

            scriptSB.AppendLine(@$"
namespace {packageName}
{{
    public sealed partial class {SpecialExcelCfg.ErrorCodeScriptName}
    {{
{fieldStr}
    }}
}}
");

            await sw.WriteAsync(scriptSB.ToString());
            sw.Flush();
        }

        private async Task WriteSingleExcelScript(string scriptName, string outputPath, string fieldStr)
        {
            using StreamWriter sw = new StreamWriter(outputPath);
            StringBuilder scriptSB = new();

            scriptSB.AppendLine(@$"
using System.Collections.Generic;

namespace {GeneralCfg.ProtoMetaPackageName}
{{
    public static class {scriptName}
    {{
{fieldStr}
    }}
}}
");

            await sw.WriteAsync(scriptSB.ToString());
            sw.Flush();
        }
    }
}
