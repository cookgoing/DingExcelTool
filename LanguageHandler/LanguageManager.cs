using System.Text;

namespace DingExcelTool.LanguageHandler;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.Concurrent;
using Azure;
using Azure.AI.Translation.Text;
using ClosedXML.Excel;
using ConcurrentCollections;
using Data;
using Configure;
using ExcelHandler;
using ScriptHandler;
using Utils;

internal class LanguageManager : Singleton<LanguageManager>
{
    public DataWraper? Data => ExcelManager.Instance.Data;
    public CommonExcelLanguageHandler CommonExcelLanguageHandler { get; private set; }
    public SingleExcelLanguageHandler SingleExcelLanguageHandler { get; private set; }
    public ErrorCodeExcelLanguageHandler ErrorCodeExcelLanguageHandler { get; private set; }
    protected Assembly outLanguageHandlerAssembly;
    protected ConcurrentDictionary<string, Type> typeDic = new();
    protected ConcurrentDictionary<string, object> objDic = new();
    
    private CancellationTokenSource cts;
    private ParallelOptions options;
    
    private TextTranslationClient _translationClient;
    private TextTranslationClient TranslationClient
    {
        get
        {
            if (_translationClient != null) return _translationClient;
            
            var credential = new AzureKeyCredential(Data.Language.AzureApiKey);
            return _translationClient = new TextTranslationClient(credential, new Uri(Data.Language.AzureEndpoint), Data.Language.AzureAraea);
        }
    }

    public LanguageManager()
    {
        CommonExcelLanguageHandler = new();
        SingleExcelLanguageHandler = new();
        ErrorCodeExcelLanguageHandler = new();
        
        cts = new CancellationTokenSource();
        options = new ParallelOptions
        {
            MaxDegreeOfParallelism = Environment.ProcessorCount - 1,
            CancellationToken = cts.Token
        };
    }

    public void Clear()
    {
        cts.Dispose();
    }

    public void Reset()
    {
        cts = new CancellationTokenSource();
        options.CancellationToken = cts.Token;
    }

    private async Task DynamicScriptAsync()
    {
        if (outLanguageHandlerAssembly != null) return;
        
        string[] scriptContent =
        [
            await File.ReadAllTextAsync(LanguageCfg.OuterLanguageHandlerInterfacePath),
            await File.ReadAllTextAsync(LanguageCfg.UILanguageHandlerPath),
            await File.ReadAllTextAsync(LanguageCfg.CodeLanguageHandlerPath)
        ];
        
        SyntaxTree[] syntaxTrees = new SyntaxTree[scriptContent.Length];
        for (int i = 0; i < scriptContent.Length; ++i) syntaxTrees[i] = CSharpSyntaxTree.ParseText(scriptContent[i]);
        
        var tpa = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")).Split(Path.PathSeparator);
        var needed = tpa.Where(p => new[]
                {
                    "System.Private.CoreLib.dll",
                    "System.Runtime.dll",
                    "System.Runtime.Extensions.dll",
                    "System.IO.dll",
                    "System.Collections.Concurrent.dll",
                    "System.Linq.dll",
                    "System.Text.RegularExpressions.dll",
                    "System.Private.Xml.Linq.dll",
                    "System.Private.Xml.dll"
                }.Contains(Path.GetFileName(p))).Select(p => MetadataReference.CreateFromFile(p)).ToList();
        
        CSharpCompilation compilation = CSharpCompilation.Create(
            "outLanguageHandlerAssembly",
            syntaxTrees: syntaxTrees,
            references: needed,
            options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        using var memoryStream = new MemoryStream();
        EmitResult result = compilation.Emit(memoryStream);
        
        if (!result.Success)
        {
            StringBuilder sb = new();
            foreach (var diagnostic in result.Diagnostics) sb.AppendLine(diagnostic.ToString());

            throw new Exception(sb.ToString());
        }
        
        memoryStream.Seek(0, SeekOrigin.Begin);
        outLanguageHandlerAssembly = Assembly.Load(memoryStream.ToArray());
    }
    public (Type type, object obj) GetTypeObj(string scriptName)
    {
        // string fullScriptName = $"{scriptName}";
        if (!typeDic.TryGetValue(scriptName, out Type type))
        {
            type = outLanguageHandlerAssembly.GetType(scriptName) ?? throw new Exception($"[GenerateTypeObj] proto生成的C#程序集不存在 这个类型：{scriptName}");
            typeDic.TryAdd(scriptName, type);
        }
        if (!objDic.TryGetValue(scriptName, out object obj))
        {
            obj = Activator.CreateInstance(type) ?? throw new Exception($"[GenerateTypeObj] 无法生成实例：type: {type}");
            objDic.TryAdd(scriptName, obj);
        }

        return (type, obj);
    }
    
    public async Task<bool> ExcelGenerate(bool autoTranslation)
    {
        await DynamicScriptAsync();
        bool result = await ExcelManager.Instance.GenerateExcelHeadInfo();
        if (!result) return false;
        
        var (uiType, uiObj) = GetTypeObj(LanguageCfg.UILanguageHandlerClassName);
        var (codeType, codeObj) = GetTypeObj(LanguageCfg.CodeLanguageHandlerClassName);
        PropertyInfo uiPropertyInfo = uiType.GetProperty("Suffix", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.UILanguageHandlerClassName}]中，没有这个字段：Suffix;");
        PropertyInfo codePropertyInfo = codeType.GetProperty("Suffix", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.CodeLanguageHandlerClassName}]中，没有这个字段：Suffix;");
        var uiObjProperty = uiPropertyInfo.GetValue(uiObj);
        var codeObjProperty = codePropertyInfo.GetValue(codeObj);
        
        string[] excelPathArr = Directory.GetFiles(Data.ExcelInputRootDir, "*.xlsx", SearchOption.AllDirectories);
        string[] uiPathArr = Directory.GetFiles(Data.Language.UIRootDir, "*" + uiObjProperty, SearchOption.AllDirectories);
        string[] codePathArr = Directory.GetFiles(Data.Language.ScriptRootDir, "*" + codeObjProperty, SearchOption.AllDirectories);
        ConcurrentHashSet<string> languages = new (), images = new();
        await Parallel.ForEachAsync(excelPathArr, options, async (excelFilePath, token) =>
        {
            try
            {
                if (token.IsCancellationRequested) return;
                
                string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
                if (excelFileName == SpecialExcelCfg.EnumExcelName) return;

                LogMessageHandler.AddInfo($"【收集多语言-表】:{excelFilePath}");
                HashSet<string> languageHash, imageHash;
                if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) (languageHash, imageHash) = await SingleExcelLanguageHandler.GetLanguagesAsync(excelFilePath);
                else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) (languageHash, imageHash) = await ErrorCodeExcelLanguageHandler.GetLanguagesAsync(excelFilePath);
                else (languageHash, imageHash) = await CommonExcelLanguageHandler.GetLanguagesAsync(excelFilePath);
                
                UnionWith(languages, languageHash);
                UnionWith(images, imageHash);
            }
            catch (Exception e)
            {
                result = false;
                LogMessageHandler.LogException(e);
                cts.Cancel();
            }
        });

        MethodInfo uiMethodInfo = uiType.GetMethod("GetLanguagesAsync", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.UILanguageHandlerClassName}]中，没有这个方法：GetLanguagesAsync;");
        MethodInfo codeMethodInfo = codeType.GetMethod("GetLanguagesAsync", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.CodeLanguageHandlerClassName}]中，没有这个方法：GetLanguagesAsync;");
        await Parallel.ForEachAsync(uiPathArr, options, async (uiFilePath, token) =>
        {
            try
            {
                if (token.IsCancellationRequested) return;
                
                LogMessageHandler.AddInfo($"【收集多语言-ui】:{uiFilePath}");
                Task<(HashSet<string>, HashSet<string>)> task = (Task<(HashSet<string>, HashSet<string>)>)uiMethodInfo.Invoke(uiObj, [uiFilePath, LogMessageHandler.AddError]);
                var (languageHash, imageHash) = await task;
                UnionWith(languages, languageHash);
                UnionWith(images, imageHash);
            }
            catch (Exception e)
            {
                result = false;
                LogMessageHandler.LogException(e);
                cts.Cancel();
            }
        });
        
        await Parallel.ForEachAsync(codePathArr, options, async (codeFilePath, token) =>
        {
            try
            {
                if (token.IsCancellationRequested) return;
                
                LogMessageHandler.AddInfo($"【收集多语言-code】:{codeFilePath}");
                Task<(HashSet<string>, HashSet<string>)> task = (Task<(HashSet<string>, HashSet<string>)>)codeMethodInfo.Invoke(codeObj, [codeFilePath, LogMessageHandler.AddError]);
                var (languageHash, imageHash) = await task;
                UnionWith(languages, languageHash);
                UnionWith(images, imageHash);
            }
            catch (Exception e)
            {
                result = false;
                LogMessageHandler.LogException(e);
                cts.Cancel();
            }
        });

        BackupLanguageExcel();
        GenerateLanguageExcel(languages, true, autoTranslation);
        GenerateLanguageExcel(images, false, autoTranslation);
        return result;
    }

    public async Task<bool> LanguageReplaceAndExportExcel()
    {
        await DynamicScriptAsync();
        bool result = await ExcelManager.Instance.GenerateExcelHeadInfo();
        result &= await ExcelManager.Instance.GenerateProtoMeta();
        result &= await ExcelManager.Instance.GenerateProtoScript();
        
        ConcurrentDictionary<string, int> languageDic = ReadLanguageExcel(true);
        ConcurrentDictionary<string, int> imageDic = ReadLanguageExcel(false);
        result &= await GenerateProtoData(languageDic, imageDic);
        result &= await GenerateExcelScript(languageDic, imageDic);

        var (uiType, uiObj) = GetTypeObj(LanguageCfg.UILanguageHandlerClassName);
        var (codeType, codeObj) = GetTypeObj(LanguageCfg.CodeLanguageHandlerClassName);
        PropertyInfo uiPropertyInfo = uiType.GetProperty("Suffix", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.UILanguageHandlerClassName}]中，没有这个字段：Suffix;");
        PropertyInfo codePropertyInfo = codeType.GetProperty("Suffix", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.CodeLanguageHandlerClassName}]中，没有这个字段：Suffix;");
        var uiObjProperty = uiPropertyInfo.GetValue(uiObj);
        var codeObjProperty = codePropertyInfo.GetValue(codeObj);
        string[] uiPathArr = Directory.GetFiles(Data.Language.UIRootDir, "*" + uiObjProperty, SearchOption.AllDirectories);
        string[] codePathArr = Directory.GetFiles(Data.Language.ScriptRootDir, "*" + codeObjProperty, SearchOption.AllDirectories);
        if (result)
        {
            MethodInfo uiMethodInfo = uiType.GetMethod("LanguageReplaceAsync", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.UILanguageHandlerClassName}]中，没有这个方法：LanguageReplaceAsync;");
            await Parallel.ForEachAsync(uiPathArr, options, async (uiFilePath, token) =>
            {
                try
                {
                    if (token.IsCancellationRequested) return;
                    LogMessageHandler.AddInfo($"【替换多语言-ui】:{uiFilePath}");
                    await (Task)uiMethodInfo.Invoke(uiObj, [uiFilePath, languageDic, imageDic, LogMessageHandler.AddError]);
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.LogException(e);
                    cts.Cancel();
                }
            });
        }

        if (result)
        {
            MethodInfo codeMethodInfo = codeType.GetMethod("LanguageReplaceAsync", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.CodeLanguageHandlerClassName}]中，没有这个方法：LanguageReplaceAsync;");
            await Parallel.ForEachAsync(codePathArr, options, async (codeFilePath, token) =>
            {
                try
                {
                    if (token.IsCancellationRequested) return;
                    LogMessageHandler.AddInfo($"【替换多语言-code】:{codeFilePath}");
                    await (Task)codeMethodInfo.Invoke(codeObj, [codeFilePath, languageDic, imageDic, LogMessageHandler.AddError]);
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.LogException(e);
                    cts.Cancel();
                }
            });
        }
        return result;
    }
    private async Task<bool> GenerateProtoData(ConcurrentDictionary<string, int> languageDic, ConcurrentDictionary<string, int> imageDic)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        bool result = true;
        string[] excelPathArr = Directory.GetFiles(Data.ExcelInputRootDir, "*.xlsx", SearchOption.AllDirectories);
        bool turnonClient = Data.OutputClient;
        bool turnonServer = Data.OutputServer;

        if (turnonClient)
        {
            string clientProtoDataOutputDir = Data.ClientOutputInfo.ProtoDataOutputDir;
            ScriptTypeEn clientScriptType = Data.ClientOutputInfo.ScriptType;

            await Task.Run(async () =>
            {
                LogMessageHandler.AddInfo($"【客户端代码动态编译】");
                string clientProtoScriptOutputDir = Data.ClientOutputInfo.ProtoScriptOutputDir;
                IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(clientScriptType);
                string[] scriptPathArr = Directory.GetFiles(clientProtoScriptOutputDir, $"*{scriptHandler.Suffix}", SearchOption.AllDirectories);
                string[] scriptContent = new string[scriptPathArr.Length];
                for (int i = 0; i < scriptPathArr.Length; ++i) scriptContent[i] = await File.ReadAllTextAsync(scriptPathArr[i]);
                
                scriptHandler.DynamicCompile(scriptContent);
            });

            IScriptLanguageHandler scriptLanguageHandler = LanguageUtils.GetScriptLanguageHandler(clientScriptType);
            string commonProtoScriptPath = Path.Combine(ExcelManager.Instance.Data.ClientOutputInfo.ProtoScriptOutputDir, $"{CommonExcelCfg.ProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix.Replace(".pbmeta", scriptLanguageHandler.Suffix)}");
            string errorcodeProtoScriptPath = Path.Combine(ExcelManager.Instance.Data.ClientOutputInfo.ProtoScriptOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix.Replace(".pbmeta", scriptLanguageHandler.Suffix)}");
            StringBuilder commonProtoScriptContentSB = new(await File.ReadAllTextAsync(commonProtoScriptPath));
            StringBuilder errorCodeProtoScriptContentSB = new(await File.ReadAllTextAsync(errorcodeProtoScriptPath));
            await Parallel.ForEachAsync(excelPathArr, options, async (excelFilePath, token) =>
            {
                try
                {
                    if (token.IsCancellationRequested) return;

                    LogMessageHandler.AddInfo($"【客户端生成proto数据】:{excelFilePath}");
                    string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
                    if (excelFileName == SpecialExcelCfg.EnumExcelName) await ExcelManager.Instance.EnumExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                    else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName)
                    {
                        if (ErrorCodeExcelLanguageHandler.IsSkip(excelFilePath)) await ExcelManager.Instance.ErrorCodeExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else await ErrorCodeExcelLanguageHandler.LanguageReplaceAsync(excelFilePath, clientProtoDataOutputDir, true, clientScriptType, languageDic, imageDic, errorCodeProtoScriptContentSB);
                    }
                    else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) await ExcelManager.Instance.SingleExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                    else
                    {
                        if (CommonExcelLanguageHandler.IsSkip(excelFilePath)) await ExcelManager.Instance.CommonExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else await CommonExcelLanguageHandler.LanguageReplaceAsync(excelFilePath, clientProtoDataOutputDir, true, clientScriptType, languageDic, imageDic, commonProtoScriptContentSB);
                    }
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.LogException(e);
                    cts.Cancel();
                }
            });
            await File.WriteAllTextAsync(commonProtoScriptPath, commonProtoScriptContentSB.ToString());
            await File.WriteAllTextAsync(errorcodeProtoScriptPath, errorCodeProtoScriptContentSB.ToString());
        }

        if (!result) goto GotoResult;

        if (turnonServer)
        {
            string serverProtoDataOutputDir = Data.ServerOutputInfo.ProtoDataOutputDir;
            ScriptTypeEn serverScriptType = Data.ServerOutputInfo.ScriptType;

            await Task.Run(async () =>
            {
                LogMessageHandler.AddInfo($"【服务器代码动态编译】");
                string serverProtoScriptOutputFile = Data.ServerOutputInfo.ProtoScriptOutputDir;
                IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(serverScriptType);
                string[] scriptPathArr = Directory.GetFiles(serverProtoScriptOutputFile, $"*{scriptHandler.Suffix}", SearchOption.AllDirectories);
                string[] scriptContent = new string[scriptPathArr.Length];
                for (int i = 0; i < scriptPathArr.Length; ++i) scriptContent[i] = await File.ReadAllTextAsync(scriptPathArr[i]);
                
                scriptHandler.DynamicCompile(scriptContent);
            });

            await Parallel.ForEachAsync(excelPathArr, options, async (excelFilePath, token) =>
            {
                try
                {
                    if (token.IsCancellationRequested) return;

                    LogMessageHandler.AddInfo($"【服务器生成proto数据】:{excelFilePath}");
                    string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
                    if (excelFileName == SpecialExcelCfg.EnumExcelName) await ExcelManager.Instance.EnumExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                    else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) await ExcelManager.Instance.ErrorCodeExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                    else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) await ExcelManager.Instance.SingleExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                    else await ExcelManager.Instance.CommonExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.LogException(e);
                    cts.Cancel();
                }
            });
        }
        
    GotoResult:
        stopwatch.Stop();
        LogMessageHandler.AddInfo($"【动态编译 + 生成proto数据 完成】花费时间：{stopwatch.ElapsedMilliseconds / 1000f}s");
        return result;
    }
    private async Task<bool> GenerateExcelScript(ConcurrentDictionary<string, int> languageDic, ConcurrentDictionary<string, int> imageDic)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        bool result = true;
        string[] excelPathArr = Directory.GetFiles(Data.ExcelInputRootDir, "*.xlsx", SearchOption.AllDirectories);
        bool turnonClient = Data.OutputClient;
        bool turnonServer = Data.OutputServer;

        await Parallel.ForEachAsync(excelPathArr, options, async (excelFilePath, token) =>
        {
            try
            {
                if (token.IsCancellationRequested) return;

                LogMessageHandler.AddInfo($"【生成Excel 脚本】:{excelFilePath}");
                string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);

                if (turnonClient)
                {
                    string clientExcelScriptOutputDir = Data.ClientOutputInfo.ExcelScriptOutputDir;
                    ScriptTypeEn clientScriptType = Data.ClientOutputInfo.ScriptType;

                    if (excelFileName == SpecialExcelCfg.EnumExcelName) await ExcelManager.Instance.EnumExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                    else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) await ExcelManager.Instance.ErrorCodeExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                    else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix))
                    {
                        if (SingleExcelLanguageHandler.IsSkip(excelFilePath)) await ExcelManager.Instance.SingleExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                        else await SingleExcelLanguageHandler.LanguageReplaceAsync(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType, languageDic, imageDic);
                    }
                    else await ExcelManager.Instance.CommonExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                }
                if (turnonServer)
                {
                    string serverExcelScriptOutputDir = Data.ServerOutputInfo.ExcelScriptOutputDir;
                    ScriptTypeEn serverScriptType = Data.ServerOutputInfo.ScriptType;

                    if (excelFileName == SpecialExcelCfg.EnumExcelName) await ExcelManager.Instance.EnumExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                    else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) await ExcelManager.Instance.ErrorCodeExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                    else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) await ExcelManager.Instance.SingleExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                    else await ExcelManager.Instance.CommonExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                }
            }
            catch (Exception e)
            {
                result = false;
                LogMessageHandler.LogException(e);
                cts.Cancel();
            }
        });

        stopwatch.Stop();
        LogMessageHandler.AddInfo($"【生成Excel脚本 完成】花费时间：{stopwatch.ElapsedMilliseconds / 1000f}s");
        return result;
    }

    public async Task<bool> LanguageRevertAndExportExcel()
    {
        await DynamicScriptAsync();
        bool result = await ExcelManager.Instance.GenerateExcelHeadInfo();
        result &= await ExcelManager.Instance.GenerateProtoMeta();
        result &= await ExcelManager.Instance.GenerateProtoScript();
        result &= await ExcelManager.Instance.GenerateProtoData();
        result &= await ExcelManager.Instance.GenerateExcelScript();
        
        ConcurrentDictionary<string, int> languageDic = ReadLanguageExcel(true);
        ConcurrentDictionary<string, int> imageDic = ReadLanguageExcel(false);
        ConcurrentDictionary<int, string> languageDicInverse = new (languageDic.Select(kv => new KeyValuePair<int,string>(kv.Value, kv.Key)));
        ConcurrentDictionary<int, string> imageDicInverse = new (imageDic.Select(kv => new KeyValuePair<int,string>(kv.Value, kv.Key)));
        var (uiType, uiObj) = GetTypeObj(LanguageCfg.UILanguageHandlerClassName);
        var (codeType, codeObj) = GetTypeObj(LanguageCfg.CodeLanguageHandlerClassName);
        PropertyInfo uiPropertyInfo = uiType.GetProperty("Suffix", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.UILanguageHandlerClassName}]中，没有这个字段：Suffix;");
        PropertyInfo codePropertyInfo = codeType.GetProperty("Suffix", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.CodeLanguageHandlerClassName}]中，没有这个字段：Suffix;");
        var uiObjProperty = uiPropertyInfo.GetValue(uiObj);
        var codeObjProperty = codePropertyInfo.GetValue(codeObj);
        string[] uiPathArr = Directory.GetFiles(Data.Language.UIRootDir, "*" + uiObjProperty, SearchOption.AllDirectories);
        string[] codePathArr = Directory.GetFiles(Data.Language.ScriptRootDir, "*" + codeObjProperty, SearchOption.AllDirectories);
        if (result)
        {
            MethodInfo uiMethodInfo = uiType.GetMethod("LanguageRevertAsync", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.UILanguageHandlerClassName}]中，没有这个方法：LanguageRevertAsync;");
            await Parallel.ForEachAsync(uiPathArr, options, async (uiFilePath, token) =>
            {
                try
                {
                    if (token.IsCancellationRequested) return;
                    LogMessageHandler.AddInfo($"【还原多语言-ui】:{uiFilePath}");
                    await (Task)uiMethodInfo.Invoke(uiObj, [uiFilePath, languageDicInverse, imageDicInverse, LogMessageHandler.AddError]);
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.LogException(e);
                    cts.Cancel();
                }
            });
        }

        if (result)
        {
            MethodInfo codeMethodInfo = codeType.GetMethod("LanguageRevertAsync", BindingFlags.Public | BindingFlags.Instance) ?? throw new Exception($"[ExcelGenerate] C#脚本[{LanguageCfg.CodeLanguageHandlerClassName}]中，没有这个方法：LanguageRevertAsync;");
            await Parallel.ForEachAsync(codePathArr, options, async (codeFilePath, token) =>
            {
                try
                {
                    if (token.IsCancellationRequested) return;
                    LogMessageHandler.AddInfo($"【还原多语言-code】:{codeFilePath}");
                    await (Task)codeMethodInfo.Invoke(codeObj, [codeFilePath, languageDicInverse, imageDicInverse, LogMessageHandler.AddError]);
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.LogException(e);
                    cts.Cancel();
                }
            });
        }
        return result;
    }

    private void BackupLanguageExcel()
    {
        string languageExcelDir = Data.Language.LanguageExcelDir;
        if (!Directory.Exists(languageExcelDir)) Directory.CreateDirectory(languageExcelDir);
        else
        {
            if (!Directory.Exists(Data.Language.LanguageBackupDir)) Directory.CreateDirectory(Data.Language.LanguageBackupDir);
            ExcelUtil.CopyDirectoryWithAutoRename(languageExcelDir, Data.Language.LanguageBackupDir);
        }
    }
    private async void GenerateLanguageExcel(ConcurrentHashSet<string> hashSet, bool isLanguage, bool autoTranslation)
    {
        if (hashSet == null || hashSet.Count == 0) return;
        
        string languageExcelDir = Data.Language.LanguageExcelDir;
        string excelFileName = isLanguage ? LanguageCfg.LanguageExcelName : LanguageCfg.ImageExcelName;
        string sheetName = isLanguage ? LanguageCfg.LanguageExcelSheetName : LanguageCfg.ImageExcelSheetName;
        var excelNameArr = isLanguage ? LanguageCfg.LanguageExcelNameArr : LanguageCfg.ImageExcelNameArr;
        var excelHeadDic = isLanguage ? LanguageCfg.LanguageExcelHeadDic : LanguageCfg.ImageExcelHeadDic;
        foreach (LanguageType lang in Enum.GetValues<LanguageType>())
        {
            string excelName = string.Format(excelFileName, lang);
            string excelFilePath = Path.Combine(languageExcelDir, excelName);
            bool isOriginalLanguage = lang == Data.Language.SourceLanguage;
            using XLWorkbook workbook = new XLWorkbook();
            IXLWorksheet ws = workbook.Worksheets.Add(sheetName);

            LogMessageHandler.AddInfo($"【生成多语言表】 excelName: {excelName}");
            if (isOriginalLanguage)
            {
                string hashIdName = excelNameArr[0];
                string contentName = excelNameArr[2];
                ws.Cell("A1").Value = $"#{HeadType.name}";
                ws.Cell("B1").Value = hashIdName;
                ws.Cell("C1").Value = contentName;

                ws.Cell("A2").Value = $"#{HeadType.type}";
                ws.Cell("B2").Value = excelHeadDic[hashIdName].type;
                ws.Cell("C2").Value = excelHeadDic[contentName].type;

                ws.Cell("A3").Value = $"#{HeadType.platform}";
                ws.Cell("B3").Value = excelHeadDic[hashIdName].platform;
                ws.Cell("C3").Value = excelHeadDic[contentName].platform;

                ws.Cell("A4").Value = $"#{HeadType.comment}";
                ws.Cell("B4").Value = excelHeadDic[hashIdName].comment;
                ws.Cell("C4").Value = excelHeadDic[contentName].comment;

                int row = 5;
                foreach (var src in hashSet)
                {
                    int id = LanguageUtils.Hash32Bytes(src);
                    ws.Cell(row, 2).Value = id;
                    ws.Cell(row, 3).Value = src;
                    row++;
                }
            }
            else
            {
                string hashIdName = excelNameArr[0];
                string sourceName = excelNameArr[1];
                string contentName = excelNameArr[2];
                ws.Cell("A1").Value = $"#{HeadType.name}";
                ws.Cell("B1").Value = hashIdName;
                ws.Cell("C1").Value = sourceName;
                ws.Cell("D1").Value = contentName;

                ws.Cell("A2").Value = $"#{HeadType.type}";
                ws.Cell("B2").Value = excelHeadDic[hashIdName].type;
                ws.Cell("C2").Value = excelHeadDic[sourceName].type;
                ws.Cell("D2").Value = excelHeadDic[contentName].type;

                ws.Cell("A3").Value = $"#{HeadType.platform}";
                ws.Cell("B3").Value = excelHeadDic[hashIdName].platform;
                ws.Cell("C3").Value = excelHeadDic[sourceName].platform;
                ws.Cell("D3").Value = excelHeadDic[contentName].platform;

                ws.Cell("A4").Value = $"#{HeadType.comment}";
                ws.Cell("B4").Value = excelHeadDic[hashIdName].comment;
                ws.Cell("C4").Value = excelHeadDic[sourceName].comment;
                ws.Cell("D4").Value = excelHeadDic[contentName].comment;

                int row = 5;
                foreach (var src in hashSet)
                {
                    int id = LanguageUtils.Hash32Bytes(src);
                    string text = string.Empty;
                    if (isLanguage && autoTranslation) text = await TranslateAsync(src, Data.Language.SourceLanguage, lang);

                    ws.Cell(row, 2).Value = id;
                    ws.Cell(row, 3).Value = src;
                    ws.Cell(row, 4).Value = text;
                    row++;
                }
            }
            var header = ws.Range("A1:D1");
            header.Style.Font.Bold = true;
            header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Columns().AdjustToContents();

            workbook.SaveAs(excelFilePath);
        }
    }

    private ConcurrentDictionary<string, int> ReadLanguageExcel(bool isLanguage)
    {
        ConcurrentDictionary<string, int> result = new();
        string languageExcelDir = Data.Language.LanguageExcelDir;
        string excelFileName = isLanguage ? LanguageCfg.LanguageExcelName : LanguageCfg.ImageExcelName;
        string excelName = string.Format(excelFileName, Data.Language.SourceLanguage);
        string languageExcelFilePath = Path.Combine(languageExcelDir, excelName);
        using XLWorkbook wb = new XLWorkbook(languageExcelFilePath);
        IXLWorksheet sheet = wb.Worksheets.Worksheet(1);

        foreach (IXLRow row in sheet.RowsUsed())
        {
            string typeName = row.Cell(1).GetString();
            bool isTpyeRow = typeName.StartsWith('#');
            if (isTpyeRow) continue;

            IXLCell hashIDCell = row.Cell(2);
            IXLCell contentCell = row.Cell(3);

            if (!int.TryParse(hashIDCell.GetString().Trim(), out int hashId)) throw new Exception($"{hashIDCell.GetString()} 不是数字 ID");
            if (!result.TryAdd(contentCell.GetString(), hashId)) throw new Exception($"有重复的 多语言字段：{contentCell.GetString()}");
        }

        return result;
    }

    private void UnionWith<T>(ConcurrentHashSet<T> hashSet, IEnumerable<T> other)
    {
        foreach (var item in other) 
            hashSet.Add(item); 
    }
    
    private string AzureLanguageType(LanguageType value)
    => value switch
    {
        LanguageType.zh_CN => "zh-Hans",
        LanguageType.zh_TW => "zh-Hant",
        LanguageType.ja => "ja",
        LanguageType.ko => "ko",
        LanguageType.en => "en",
        LanguageType.fr => "fr",
        LanguageType.de => "de",
        LanguageType.ru => "ru",
        LanguageType.es => "es",
        LanguageType.pt => "pt",
        _ => "zh-Hans",
    };
    private async Task<string> TranslateAsync(string inputText, LanguageType from, LanguageType to)
    {
        string fromLang = AzureLanguageType(from);
        string toLang = AzureLanguageType(to);
        try
        {
            var response = await TranslationClient.TranslateAsync(
                sourceLanguage: fromLang,
                targetLanguages: new[] { toLang },
                content: new[] { inputText }
            );

            return response.Value[0].Translations[0].Text;
        }
        catch (Exception ex)
        {
            LogMessageHandler.LogException(ex);
        }
        
        return null;
    }
}