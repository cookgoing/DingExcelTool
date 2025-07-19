namespace DingExcelTool.LanguageHandler;

using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.Concurrent;
using ClosedXML.Excel;
using Data;
using Utils;
using Configure;
using ExcelHandler;
using ScriptHandler;

internal class ErrorCodeExcelLanguageHandler : CommonExcelLanguageHandler
{
    public override bool IsSkip(string inputPath)
    {
        if (inputPath.StartsWith(ExcelManager.Instance.Data.Language.LanguageExcelDir)) return true;

        string excelName = Path.GetFileNameWithoutExtension(inputPath);
        if (excelName != SpecialExcelCfg.ErrorCodeExcelName) return true;

        return false;
    }

    public override async Task<(HashSet<string> languageHash, HashSet<string> imageHash)> GetLanguagesAsync(string inputPath)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[GetLanguagesAsync] 表路径不存在：{inputPath}");

        HashSet<string> languageHash = new(), imageHash = new();
        string excelFileName = Path.GetFileNameWithoutExtension(inputPath);
        using XLWorkbook wb = new(inputPath);
        int sheetCount = wb.Worksheets.Count;
        ErrorCodeExcelHandler errorCodeExcelHandler = ExcelManager.Instance.ErrorCodeExcelHandler;
        foreach (IXLWorksheet sheet in wb.Worksheets)
        {
            string messageName = sheetCount == 1 ? $"{excelFileName}" : $"{excelFileName}_{sheet.Name}";
            if (!errorCodeExcelHandler.HeadInfoDic.TryGetValue(messageName, out ExcelHeadInfo headInfo)) continue;

            foreach (IXLRow row in sheet.RowsUsed())
            {
                string typeName = row.Cell(1).GetString();
                bool isTpyeRow = typeName.StartsWith('#');
                if (isTpyeRow) continue;

                foreach (IXLCell cell in row.Cells(false))
                {
                    int columnIdx = cell.Address.ColumnNumber;
                    if (columnIdx == 1) continue;
                    if (ExcelUtil.IsRearMergedCell(cell)) continue;

                    int fieldIdx = headInfo.GetFieldIdx(columnIdx);
                    if (fieldIdx == -1) continue;

                    string columnContent = cell.GetString().Trim();
                    ExcelFieldInfo fieldInfo = headInfo.Fields[fieldIdx];
                    if (ExcelUtil.IsMapType(fieldInfo.Type))
                    {
                        int relativeColumnIdx = columnIdx - fieldInfo.StartColumnIdx;
                        bool isKey = relativeColumnIdx % 2 == 0;
                        switch (isKey)
                        {
                            case true when fieldInfo.LocalizationTxt.k:
                            case false when fieldInfo.LocalizationTxt.v:
                                languageHash.Add(columnContent);
                                break;
                            case true when fieldInfo.LocalizationImg.k:
                            case false when fieldInfo.LocalizationImg.v:
                                imageHash.Add(columnContent);
                                break;
                        }
                    }
                    else if (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v)
                    {
                        languageHash.Add(columnContent);
                    }
                    else if (fieldInfo.LocalizationImg.k || fieldInfo.LocalizationImg.v)
                    {
                        imageHash.Add(columnContent);
                    }
                }
            }
        }

        return (languageHash, imageHash);
    }

    public override async Task LanguageReplaceAsync(string inputPath, string protoDataOutputDir, bool isClient, ScriptTypeEn scriptType, ConcurrentDictionary<string, int> languageDic, ConcurrentDictionary<string, int> imageDic, params object[] arg)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[LanguageReplaceAsync] 表路径不存在：{inputPath}");
        if (languageDic == null)  throw new ArgumentNullException(nameof(languageDic));
        if (arg.Length < 1 || arg[0] is not StringBuilder errorcodeProtoScriptContentSB) throw new Exception($"[LanguageReplaceAsync] no commonProtoScriptContent");
        string errorcodeProtoScriptContent = errorcodeProtoScriptContentSB.ToString();
        if (string.IsNullOrEmpty(protoDataOutputDir))
        {
            LogMessageHandler.AddWarn($"[LanguageReplaceAsync] 不存在 proto data 的输出路径，将不会执行输出操作");
            return;
        }

        IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(scriptType);
        PlatformType platform = isClient ? PlatformType.Client : PlatformType.Server;
        string excelFileName = Path.GetFileNameWithoutExtension(inputPath);
        using XLWorkbook wb = new(inputPath);
        int sheetCount = wb.Worksheets.Count;
        ErrorCodeExcelHandler errorCodeExcelHandler = ExcelManager.Instance.ErrorCodeExcelHandler;
        foreach (IXLWorksheet sheet in wb.Worksheets)
        {
            string messageName = sheetCount == 1 ? $"{excelFileName}" : $"{excelFileName}_{sheet.Name}";
            if (!errorCodeExcelHandler.HeadInfoDic.TryGetValue(messageName, out ExcelHeadInfo headInfo)) throw new Exception($"[LanguageReplaceAsync] 表: {messageName} 没有 headInfo");

            AddProtoDataInScriptObj(scriptHandler, platform, sheet, SpecialExcelCfg.ErrorCodeProtoMessageName, headInfo, languageDic, imageDic);
        }
        
        string protoDataOutputFile = Path.Combine(protoDataOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMessageName}{GeneralCfg.ProtoDataFileSuffix}");
        scriptHandler.SerializeObjInProto($"{SpecialExcelCfg.ErrorCodeProtoMessageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}", protoDataOutputFile);

        IScriptLanguageHandler scriptLanguageHandler = LanguageUtils.GetScriptLanguageHandler(scriptType);
        string languageReplaceFucName = ExcelManager.Instance.Data.Language.LanguageReplaceMethod;
        string imageReplaceFucName = ExcelManager.Instance.Data.Language.ImageReplaceMethod;
        foreach (var kv in SpecialExcelCfg.ErrorCodeFixedField)
        {
            string name = kv.Key;
            string typeName = kv.Value;
            bool isLocalizationTxt = ExcelUtil.IsTypeLocalizationTxt(typeName);
            bool isLocalizationImg = ExcelUtil.IsTypeLocalizationImg(typeName);
            if (!isLocalizationTxt && !isLocalizationImg) continue;

            typeName = ExcelUtil.ClearTypeSymbol(typeName);
            string fieldName = $"{NameConverter.ConvertToCamelCase(name)}_";
            if (ExcelUtil.IsMapType(typeName))
            {
                string innerTypes = kv.Value.Substring(4, kv.Value.Length - 5);
                string[] keyValue = innerTypes.Split(',');
                bool kLocalizationImg = ExcelUtil.IsTypeLocalizationImg(keyValue[0]);
                bool kLocalizationTxt = ExcelUtil.IsTypeLocalizationTxt(keyValue[0]);
                bool vLocalizationImg = ExcelUtil.IsTypeLocalizationImg(keyValue[1]);
                bool vLocalizationTxt = ExcelUtil.IsTypeLocalizationTxt(keyValue[1]);
                
                if (kLocalizationTxt || vLocalizationTxt) errorcodeProtoScriptContent = await scriptLanguageHandler.ReplaceMapFieldInClassProperty(errorcodeProtoScriptContent, SpecialExcelCfg.ErrorCodeProtoMessageName, fieldName, languageReplaceFucName, kLocalizationTxt, vLocalizationTxt);
                else errorcodeProtoScriptContent = await scriptLanguageHandler.ReplaceMapFieldInClassProperty(errorcodeProtoScriptContent, SpecialExcelCfg.ErrorCodeProtoMessageName, fieldName, imageReplaceFucName, kLocalizationImg, vLocalizationImg);
            }
            else if (ExcelUtil.IsArrType(typeName))
            {
                string fucName = isLocalizationTxt ? languageReplaceFucName : imageReplaceFucName;
                errorcodeProtoScriptContent = await scriptLanguageHandler.ReplaceArrFieldInClassProperty(errorcodeProtoScriptContent, SpecialExcelCfg.ErrorCodeProtoMessageName, fieldName, fucName);
            }
            else
            {
                string fucName = isLocalizationTxt ? languageReplaceFucName : imageReplaceFucName;
                errorcodeProtoScriptContent = await scriptLanguageHandler.ReplaceGetterFieldInClassProperty(errorcodeProtoScriptContent, SpecialExcelCfg.ErrorCodeProtoMessageName, fieldName, fucName);
            }
        }

        errorcodeProtoScriptContentSB.Clear().Append(errorcodeProtoScriptContent);
        await Task.CompletedTask;
    }
}