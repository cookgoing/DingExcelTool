namespace DingExcelTool.LanguageHandler;

using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Data;
using ExcelHandler;
using Configure;
using ScriptHandler;
using Utils;

internal class SingleExcelLanguageHandler : ILanguageHandler
{
    public bool IsSkip(string inputPath)
    {
        if (inputPath.StartsWith(ExcelManager.Instance.Data.Language.LanguageExcelDir)) return true;

        string excelName = Path.GetFileNameWithoutExtension(inputPath);
        if (!excelName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) return true;

        return false;
    }

    public async Task<(HashSet<string> languageHash, HashSet<string> imageHash)> GetLanguagesAsync(string inputPath)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[GetLanguagesAsync] 表路径不存在：{inputPath}");

        SingleExcelHandler singleExcelHandler = ExcelManager.Instance.SingleExcelHandler;
        string excelFileName = Path.GetFileNameWithoutExtension(inputPath);
        string excelName = excelFileName.Replace(SpecialExcelCfg.SingleExcelPrefix, string.Empty);
        using XLWorkbook wb = new XLWorkbook(inputPath);
        int sheetCount = wb.Worksheets.Count;
        HashSet<string> languageHash = new(), imageHash = new();

        foreach (IXLWorksheet sheet in wb.Worksheets)
        {
            string scriptName = sheetCount == 1 ?  $"{excelName}" : $"{excelName}_{sheet.Name}";
            if (!singleExcelHandler.HeadInfoDic.TryGetValue(scriptName, out SingleExcelHeadInfo headInfo)) throw new Exception($"[GetLanguagesAsync]. 表：{scriptName}; 没有表头信息");

            foreach (SingleExcelFieldInfo fieldInfo in headInfo.Fields)
            {
                if (!fieldInfo.LocalizationTxt.k && !fieldInfo.LocalizationTxt.v && !fieldInfo.LocalizationImg.k && !fieldInfo.LocalizationImg.v) continue;
                if (ExcelUtil.IsMapType(fieldInfo.Type))
                {
                    string[] valueStrArr = fieldInfo.Value.Split(SpecialExcelCfg.SingleArrMapSplitSymbol);
                    foreach (string kvStr in valueStrArr)
                    {
                        string[] kvStrArr = kvStr.Split(SpecialExcelCfg.SingleMapKVSplitSymbol);
                        if (fieldInfo.LocalizationTxt.k) languageHash.Add(kvStrArr[0]);
                        else if (fieldInfo.LocalizationImg.k) imageHash.Add(kvStrArr[0]);
                        
                        if (fieldInfo.LocalizationTxt.v) languageHash.Add(kvStrArr[1]);
                        else if (fieldInfo.LocalizationImg.v) imageHash.Add(kvStrArr[1]);
                    }
                }
                else if (ExcelUtil.IsArrType(fieldInfo.Type))
                {
                    string[] valueStrArr = fieldInfo.Value.Split(SpecialExcelCfg.SingleArrMapSplitSymbol);
                    var hashset = (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v) ? languageHash : imageHash;
                    foreach (string str in valueStrArr) hashset.Add(str);
                }
                else
                {
                    var hashset = (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v) ? languageHash : imageHash;
                    hashset.Add(fieldInfo.Value);
                }
            }
        }

        return (languageHash, imageHash);
    }

    public async Task LanguageReplaceAsync(string inputPath, string excelScriptOutputDir, bool isClient, ScriptTypeEn scriptType, ConcurrentDictionary<string, int> languageDic, ConcurrentDictionary<string, int> imageDic, params object[] arg)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[LanguageReplaceAsync] 表路径不存在：{inputPath}");
        if (languageDic == null)  throw new ArgumentNullException(nameof(languageDic));
        
        string excelRelativePath = ExcelUtil.GetExcelRelativePath(inputPath);
        string excelRelativeDir = Path.GetDirectoryName(excelRelativePath);
        string excelFileName = Path.GetFileNameWithoutExtension(excelRelativePath);
        string excelName = excelFileName?.Replace(SpecialExcelCfg.SingleExcelPrefix, string.Empty);
        SingleExcelHandler singleExcelHandler = ExcelManager.Instance.SingleExcelHandler;
        PlatformType platform = isClient ? PlatformType.Client : PlatformType.Server;
        string languageReplaceFucName = ExcelManager.Instance.Data.Language.LanguageReplaceMethod;
        string imageReplaceFucName = ExcelManager.Instance.Data.Language.ImageReplaceMethod;
        using XLWorkbook wb = new(inputPath);
        int sheetCount = wb.Worksheets.Count;
        
        foreach (IXLWorksheet sheet in wb.Worksheets)
        {
            string scriptName = sheetCount == 1 ? $"{excelName}" : $"{excelName}_{sheet.Name}";
            if (!singleExcelHandler.HeadInfoDic.TryGetValue(scriptName, out SingleExcelHeadInfo headInfo)) throw new Exception($"[LanguageReplaceAsync]. 表：{scriptName}; 没有表头信息");

            StringBuilder fieldSB = new(), fucSB = new("\t\tpublic static void ResetLanguage()\n\t\t{\n");
            foreach (SingleExcelFieldInfo fieldInfo in headInfo.Fields)
            {
                if ((platform & fieldInfo.Platform) == 0) continue;

                string filedValue;
                bool needNewField = false;
                if (ExcelUtil.IsArrType(fieldInfo.Type))
                {
                    StringBuilder sb = new("new[] {");
                    string baseType = fieldInfo.Type.Substring(0, fieldInfo.Type.Length - 2);
                    string[] valueStrArr = fieldInfo.Value.Split(SpecialExcelCfg.SingleArrMapSplitSymbol);
                    if (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v || fieldInfo.LocalizationImg.k || fieldInfo.LocalizationImg.v)
                    {
                        needNewField = true;
                        var dic = fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v ? languageDic : imageDic;
                        var fucName = fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v ? languageReplaceFucName : imageReplaceFucName;
                        foreach (string str in valueStrArr)
                        {
                            if (!dic.TryGetValue(str, out int hashId))
                            {
                                LogMessageHandler.AddError($"表: {scriptName}, 字段: {str} 没有生成对应的多语言文本");
                                continue;
                            }
                            string newValue = fucName.Replace(LanguageCfg.LanguageTextImageReplaceArg, hashId.ToString());
                            sb.Append(newValue).Append(',');
                        }
                    }
                    else
                    {
                        foreach (string str in valueStrArr)
                        {
                            string value = CSharpExcelHandler.Instance.ExcelType2ScriptType(baseType, str).ToString();
                            if (baseType == "string") value = $"\"{value}\"";
                            else if (baseType == "bool") value = value?.ToLower();
                            sb.Append(value).Append(','); 
                        }
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
                    foreach (string kvStr in valueStrArr)
                    {
                        string[] kvStrArr = kvStr.Split(SpecialExcelCfg.SingleMapKVSplitSymbol);
                        string kValue, vValue;
                        if (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationImg.k)
                        {
                            needNewField = true;
                            var dic = fieldInfo.LocalizationTxt.k ? languageDic : imageDic;
                            var fucName = fieldInfo.LocalizationTxt.k? languageReplaceFucName : imageReplaceFucName;
                            if (!dic.TryGetValue(kvStrArr[0], out int hashId))
                            {
                                LogMessageHandler.AddError($"表: {scriptName}, 字段: {kvStrArr[0]} 没有生成对应的多语言文本");
                                continue;
                            }
                            kValue = fucName.Replace(LanguageCfg.LanguageTextImageReplaceArg, hashId.ToString());
                        }
                        else
                        {
                            kValue = CSharpExcelHandler.Instance.ExcelType2ScriptType(kType, kvStrArr[0]).ToString();
                            if (kType == "string") kValue = $"\"{kValue}\"";
                            else if (kType == "bool") kValue = kValue?.ToLower();
                        }
                        
                        if (fieldInfo.LocalizationTxt.v || fieldInfo.LocalizationImg.v)
                        {
                            needNewField = true;
                            var dic = fieldInfo.LocalizationTxt.v ? languageDic : imageDic;
                            var fucName = fieldInfo.LocalizationTxt.v? languageReplaceFucName : imageReplaceFucName;
                            if (!dic.TryGetValue(kvStrArr[1], out int hashId))
                            {
                                LogMessageHandler.AddError($"表: {scriptName}, 字段: {kvStrArr[1]} 没有生成对应的多语言文本");
                                continue;
                            }
                            vValue = fucName.Replace(LanguageCfg.LanguageTextImageReplaceArg, hashId.ToString());
                        }
                        else
                        {
                            vValue = CSharpExcelHandler.Instance.ExcelType2ScriptType(vType, kvStrArr[1]).ToString();
                            if (vType == "string") vValue = $"\"{vValue}\"";
                            else if (vType == "bool") vValue = vValue?.ToLower();
                        }

                        sb.Append('{').Append(kValue).Append(',').Append(vValue).Append('}').Append(',');
                    }
                    sb.Remove(sb.Length - 1, 1).Append('}');
                    filedValue  = sb.ToString();
                }
                else
                {
                    if (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v || fieldInfo.LocalizationImg.k || fieldInfo.LocalizationImg.v)
                    {
                        var dic = fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v ? languageDic : imageDic;
                        var fucName = fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v ? languageReplaceFucName : imageReplaceFucName;
                        if (!dic.TryGetValue(fieldInfo.Value, out int hashId))
                        {
                            LogMessageHandler.AddError($"表: {scriptName}, 字段: {fieldInfo.Value} 没有生成对应的多语言文本");
                            continue;
                        }
                        filedValue = fucName.Replace(LanguageCfg.LanguageTextImageReplaceArg, hashId.ToString());
                    }
                    else
                    {
                        filedValue = CSharpExcelHandler.Instance.ExcelType2ScriptType(fieldInfo.Type, fieldInfo.Value).ToString();
                        if (fieldInfo.Type == "string") filedValue = $"\"{filedValue}\"";
                        else if (fieldInfo.Type == "bool") filedValue = filedValue?.ToLower();    
                    }
                }

                if (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v || fieldInfo.LocalizationImg.k || fieldInfo.LocalizationImg.v)
                {
                    string typeStr = CSharpExcelHandler.Instance.ExcelType2ScriptTypeStr(fieldInfo.Type);
                    if (needNewField)
                    {
                        string languageFieldName = $"{fieldInfo.Name}_language";
                        fieldSB.AppendLine($"\t\tprivate static {typeStr} {languageFieldName};");
                        fieldSB.Append($"\t\tpublic static {typeStr} {fieldInfo.Name} => {languageFieldName} ??= {filedValue};").AppendLine(string.IsNullOrEmpty(fieldInfo.Comment) ? null : "//" + fieldInfo.Comment);
                        fucSB.AppendLine($"\t\t\t{languageFieldName} = null;");
                    }
                    else fieldSB.Append($"\t\tpublic static {typeStr} {fieldInfo.Name} => {filedValue};").AppendLine(string.IsNullOrEmpty(fieldInfo.Comment) ? null : "//" + fieldInfo.Comment);
                }
                else fieldSB.Append($"\t\tpublic readonly static {CSharpExcelHandler.Instance.ExcelType2ScriptTypeStr(fieldInfo.Type)} {fieldInfo.Name} = {filedValue};").AppendLine(string.IsNullOrEmpty(fieldInfo.Comment) ? null : "//" + fieldInfo.Comment);
            }
            
            string outputFilePath = Path.Combine(excelScriptOutputDir, excelRelativeDir ?? "", $"{scriptName}{GeneralCfg.ExcelScriptFileSuffix(scriptType)}");
            using StreamWriter sw = new StreamWriter(outputFilePath);
            fieldSB.Remove(fieldSB.Length - 1, 1);
            fucSB.Append("\t\t}");
            StringBuilder scriptSB = new();
            scriptSB.AppendLine(@$"
using System.Collections.Generic;

namespace {GeneralCfg.ProtoMetaPackageName}
{{
    public static class {scriptName}
    {{
{fieldSB}

{fucSB}
    }}
}}
");
            await sw.WriteAsync(scriptSB.ToString());
        }
    }

    public async Task LanguageRevertAsync(string inputPath, ConcurrentDictionary<int, string> languageDic, ConcurrentDictionary<int, string> imageDic) => await Task.CompletedTask;
}