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

internal class CommonExcelLanguageHandler : ILanguageHandler
{
    public virtual bool IsSkip(string inputPath)
    {
        if (inputPath.StartsWith(ExcelManager.Instance.Data.Language.LanguageExcelDir)) return true;

        string excelName = Path.GetFileNameWithoutExtension(inputPath);
        if (excelName == SpecialExcelCfg.EnumExcelName) return true;
        if (excelName == SpecialExcelCfg.ErrorCodeExcelName) return true;
        if (excelName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) return true;

        return false;
    }

    public virtual async Task<(HashSet<string> languageHash, HashSet<string> imageHash)> GetLanguagesAsync(string inputPath)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[GetLanguagesAsync] 表路径不存在：{inputPath}");

        HashSet<string> languageHash = new(), imageHash = new();
        string excelFileName = Path.GetFileNameWithoutExtension(inputPath);
        using XLWorkbook wb = new(inputPath);
        int sheetCount = wb.Worksheets.Count;
        CommonExcelHandler commonExcelHandler = ExcelManager.Instance.CommonExcelHandler;
        
        foreach (IXLWorksheet sheet in wb.Worksheets)
        {
            string messageName = sheetCount == 1 ? $"{excelFileName}" : $"{excelFileName}_{sheet.Name}";
            if (!commonExcelHandler.HeadInfoDic.TryGetValue(messageName, out ExcelHeadInfo headInfo)) continue;

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

    public virtual async Task LanguageReplaceAsync(string inputPath, string protoDataOutputDir, bool isClient, ScriptTypeEn scriptType, ConcurrentDictionary<string, int> languageDic, ConcurrentDictionary<string, int> imageDic, params object[] arg)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[LanguageReplaceAsync] 表路径不存在：{inputPath}");
        if (languageDic == null)  throw new ArgumentNullException(nameof(languageDic));
        if (imageDic == null)  throw new ArgumentNullException(nameof(imageDic));
        if (arg.Length < 1 || arg[0] is not StringBuilder commonProtoScriptContentSB) throw new Exception($"[LanguageReplaceAsync] no commonProtoScriptContent");
        if (string.IsNullOrEmpty(protoDataOutputDir))
        {
            LogMessageHandler.AddWarn($"[LanguageReplaceAsync] 不存在 proto data 的输出路径，将不会执行输出操作");
            return;
        }
        
        string commonProtoScriptContent = commonProtoScriptContentSB.ToString();

        IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(scriptType);
        PlatformType platform = isClient ? PlatformType.Client : PlatformType.Server;
        string excelFileName = Path.GetFileNameWithoutExtension(inputPath);
        using XLWorkbook wb = new(inputPath);
        int sheetCount = wb.Worksheets.Count;
        CommonExcelHandler commonExcelHandler = ExcelManager.Instance.CommonExcelHandler;

        foreach (IXLWorksheet sheet in wb.Worksheets)
        {
            string messageName = sheetCount == 1 ? $"{excelFileName}" : $"{excelFileName}_{sheet.Name}";
            if (!commonExcelHandler.HeadInfoDic.TryGetValue(messageName, out ExcelHeadInfo headInfo)) throw new Exception($"[LanguageReplaceAsync] 表: {messageName} 没有 headInfo");

            AddProtoDataInScriptObj(scriptHandler, platform, sheet, messageName, headInfo, languageDic, imageDic);

            string protoDataOutputFile = Path.Combine(protoDataOutputDir, $"{messageName}{GeneralCfg.ProtoDataFileSuffix}");
            scriptHandler.SerializeObjInProto($"{messageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}", protoDataOutputFile);

            IScriptLanguageHandler scriptLanguageHandler = LanguageUtils.GetScriptLanguageHandler(scriptType);
            string languageReplaceFucName = ExcelManager.Instance.Data.Language.LanguageReplaceMethod;
            string imageReplaceFucName = ExcelManager.Instance.Data.Language.ImageReplaceMethod;
            foreach (ExcelFieldInfo fieldInfo in headInfo.Fields)
            {
                if (!fieldInfo.LocalizationTxt.k && !fieldInfo.LocalizationTxt.v && !fieldInfo.LocalizationImg.k && !fieldInfo.LocalizationImg.v) continue;
                
                if (ExcelUtil.IsMapType(fieldInfo.Type))
                {
                    string mapFieldName = $"{NameConverter.ConvertToCamelCase(fieldInfo.Name)}_";
                    if (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v) commonProtoScriptContent = await scriptLanguageHandler.ReplaceMapFieldInClassProperty(commonProtoScriptContent, headInfo.MessageName, mapFieldName, languageReplaceFucName, fieldInfo.LocalizationTxt.k, fieldInfo.LocalizationTxt.v);
                    else commonProtoScriptContent = await scriptLanguageHandler.ReplaceMapFieldInClassProperty(commonProtoScriptContent, headInfo.MessageName, mapFieldName, imageReplaceFucName, fieldInfo.LocalizationImg.k, fieldInfo.LocalizationImg.v);
                }
                else if (ExcelUtil.IsArrType(fieldInfo.Type))
                {
                    string arrField = $"{NameConverter.ConvertToCamelCase(fieldInfo.Name)}_";
                    string fucName = (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v) ? languageReplaceFucName : imageReplaceFucName;
                    commonProtoScriptContent = await scriptLanguageHandler.ReplaceArrFieldInClassProperty(commonProtoScriptContent, headInfo.MessageName, arrField, fucName);
                }
                else
                {
                    string field = $"{NameConverter.ConvertToCamelCase(fieldInfo.Name)}_";
                    string fucName = (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v) ? languageReplaceFucName : imageReplaceFucName;
                    commonProtoScriptContent = await scriptLanguageHandler.ReplaceGetterFieldInClassProperty(commonProtoScriptContent, headInfo.MessageName, field, fucName);
                }
            }

            commonProtoScriptContentSB.Clear().Append(commonProtoScriptContent);
        }

        await Task.CompletedTask;
    }

    public async Task LanguageRevertAsync(string inputPath, ConcurrentDictionary<int, string> languageDic, ConcurrentDictionary<int, string> imageDic) => await Task.CompletedTask;
    
    protected virtual void AddProtoDataInScriptObj(IScriptExcelHandler scriptHandler, PlatformType platform, IXLWorksheet sheet, string messageName, ExcelHeadInfo headInfo, ConcurrentDictionary<string, int> languageDic, ConcurrentDictionary<string, int> imageDic)
    {
        Dictionary<ExcelFieldInfo, HashSet<string>> singleKeyDic = new();
        HashSet<List<string>> unionKeyHash = new(new UnionKeyComparer());

        foreach (ExcelFieldInfo singleKeyField in headInfo.IndependentKey) singleKeyDic.Add(singleKeyField, new());

        foreach (IXLRow row in sheet.RowsUsed())
        {
            string typeName = row.Cell(1).GetString();
            bool isTpyeRow = typeName.StartsWith('#');
            if (isTpyeRow) continue;

            List<string> unionKeyList = new();
            foreach (IXLCell cell in row.Cells(false))
            {
                int columnIdx = cell.Address.ColumnNumber;
                if (columnIdx == 1) continue;
                if (ExcelUtil.IsRearMergedCell(cell)) continue;

                int fieldIdx = headInfo.GetFieldIdx(columnIdx);
                if (fieldIdx == -1) continue;

                string columnContent = cell.GetString().Trim();
                ExcelFieldInfo fieldInfo = headInfo.Fields[fieldIdx];
                if ((platform & fieldInfo.Platform) == 0) continue;
                if (singleKeyDic.TryGetValue(fieldInfo, out HashSet<string> singleKeyHash) && !singleKeyHash.Add(columnContent)) throw new Exception($"[GenerateProtoData].表：{messageName}-{sheet.Name} 存在了重复的独立key: {columnContent}");
                if (headInfo.UnionKey.Contains(fieldInfo)) unionKeyList.Add(columnContent);

                //LogMessageHandler.AddWarn($"[test]. messageName: {messageName}; Address: {cell.Address}; type: {fieldInfo.Type}; columnContent: {columnContent}; platform: {platform}; fieldPlatform: {fieldInfo.Platform}; isClient: {isClient}");
                if (ExcelUtil.IsMapType(fieldInfo.Type))
                {
                    int relativeColumnIdx = columnIdx - fieldInfo.StartColumnIdx;
                    bool isKey = relativeColumnIdx % 2 == 0;
                    if (!isKey) continue;

                    IXLCell nextCell = row.Cell(++columnIdx);
                    fieldIdx = headInfo.GetFieldIdx(columnIdx);
                    if (nextCell == null || fieldIdx == -1 || !ExcelUtil.IsMapType(headInfo.Fields[fieldIdx].Type)) throw new Exception($"[GenerateProtoData] 表：{messageName} 没有这个字段：{fieldInfo.Name} || 格式不合法，这里应该是一个字典值。 Address: {nextCell?.Address}");

                    string keyData = columnContent;
                    string valueData = nextCell.GetString().Trim();
                    if (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationImg.k)
                    {
                        var dic = fieldInfo.LocalizationTxt.k ? languageDic : imageDic;
                        if (dic.TryGetValue(keyData, out int hashId)) keyData = hashId.ToString();
                        else LogMessageHandler.AddError($"表: {messageName}-{sheet.Name}, 字段: {keyData}; Address: {cell.Address}; 没有生成对应的多语言文本");
                    }
                    if (fieldInfo.LocalizationTxt.v || fieldInfo.LocalizationImg.v)
                    {
                        var dic = fieldInfo.LocalizationTxt.v ? languageDic : imageDic;
                        if (dic.TryGetValue(valueData, out int hashId)) valueData = hashId.ToString();
                        else LogMessageHandler.AddError($"表: {messageName}-{sheet.Name}, 字段: {valueData}; Address: {nextCell.Address}; 没有生成对应的多语言文本");
                    }
                    scriptHandler.AddScriptMap(messageName, fieldInfo.Name, fieldInfo.Type, keyData, valueData);
                }
                else
                {
                    if (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v || fieldInfo.LocalizationImg.k || fieldInfo.LocalizationImg.v)
                    {
                        var dic = (fieldInfo.LocalizationTxt.k || fieldInfo.LocalizationTxt.v) ? languageDic : imageDic;
                        if (dic.TryGetValue(columnContent, out int hashId)) columnContent = hashId.ToString();
                        else LogMessageHandler.AddError($"表: {messageName}-{sheet.Name}, 字段: {columnContent}; Address: {cell.Address}; 没有生成对应的多语言文本");
                    }
                    
                    if (ExcelUtil.IsBaseType(fieldInfo.Type) || ExcelUtil.IsEnumType(fieldInfo.Type))
                    {
                        scriptHandler.SetScriptValue(messageName, fieldInfo.Name, fieldInfo.Type, columnContent);
                    }
                    else if (ExcelUtil.IsArrType(fieldInfo.Type))
                    {
                        scriptHandler.AddScriptList(messageName, fieldInfo.Name, fieldInfo.Type, columnContent);
                    }
                }
            }

            if (unionKeyList.Count > 0 && !unionKeyHash.Add(unionKeyList)) throw new Exception($"[GenerateProtoData]. 表：{messageName} 存在了重复的联合key: {string.Join(',', unionKeyList)}");

            scriptHandler.AddListScriptObj($"{messageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}", messageName);
        }
    }
}