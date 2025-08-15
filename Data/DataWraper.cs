namespace DingExcelTool.Data
{
    using DingExcelTool.Utils;
    
    internal class DataWraper(CustomDataInfo data)
    {
        public class PlatformOutputInfo
        {
            private CustomDataInfo.PlatformOutputInfo platformOutputInfo;

            public ScriptTypeEn ScriptType { get => platformOutputInfo.ScriptType; set => platformOutputInfo.ScriptType = value; }
            public string ProtoMetaOutputDir { get => ExcelUtil.ParsePath(platformOutputInfo.ProtoMetaOutputDir); set => platformOutputInfo.ProtoMetaOutputDir = value; }
            public string ProtoScriptOutputDir { get => ExcelUtil.ParsePath(platformOutputInfo.ProtoScriptOutputDir); set => platformOutputInfo.ProtoScriptOutputDir = value; }
            public string ProtoDataOutputDir { get => ExcelUtil.ParsePath(platformOutputInfo.ProtoDataOutputDir); set => platformOutputInfo.ProtoDataOutputDir = value; }
            public string ExcelScriptOutputDir { get => ExcelUtil.ParsePath(platformOutputInfo.ExcelScriptOutputDir); set => platformOutputInfo.ExcelScriptOutputDir = value; }
            public string ErrorCodeFrameDir { get => ExcelUtil.ParsePath(platformOutputInfo.ErrorCodeFrameDir); set => platformOutputInfo.ErrorCodeFrameDir = value; }
            public string ErrorCodeBusinessDir { get => ExcelUtil.ParsePath(platformOutputInfo.ErrorCodeBusinessDir); set => platformOutputInfo.ErrorCodeBusinessDir = value; }

            public PlatformOutputInfo(CustomDataInfo.PlatformOutputInfo platformOutputInfo) =>  this.platformOutputInfo = platformOutputInfo;
        }
        public class LanguageInfo
        {
            private CustomDataInfo.LanguageInfo languageInfo;

            public string LanguageExcelDir { get => ExcelUtil.ParsePath(languageInfo.LanguageExcelDir); set => languageInfo.LanguageExcelDir = value; }
            public string ScriptRootDir { get => ExcelUtil.ParsePath(languageInfo.ScriptRootDir); set => languageInfo.ScriptRootDir = value; }
            public string UIRootDir { get => ExcelUtil.ParsePath(languageInfo.UIRootDir); set => languageInfo.UIRootDir = value; }
            public string LanguageBackupDir { get => ExcelUtil.ParsePath(languageInfo.LanguageBackupDir); set => languageInfo.LanguageBackupDir = value; }
            public LanguageType SourceLanguage { get => languageInfo.SourceLanguage; set => languageInfo.SourceLanguage = value; }
            public string LanguageReplaceMethod { get => languageInfo.LanguageReplaceMethod; set => languageInfo.LanguageReplaceMethod = value; }
            public string ImageReplaceMethod { get => languageInfo.ImageReplaceMethod; set => languageInfo.ImageReplaceMethod = value; }
            public string AzureEndpoint { get => languageInfo.AzureEndpoint; set => languageInfo.AzureEndpoint = value; }
            public string AzureApiKey { get => languageInfo.AzureApiKey; set => languageInfo.AzureApiKey = value; }
            public string AzureAraea { get => languageInfo.AzureAraea; set => languageInfo.AzureAraea = value; }
            
            public LanguageInfo(CustomDataInfo.LanguageInfo languageInfo) =>  this.languageInfo = languageInfo;
        }

        public CustomDataInfo Data { get; private set; } = data;

        public string ExcelInputRootDir { get => ExcelUtil.ParsePath(Data.ExcelInputRootDir); set => Data.ExcelInputRootDir = value; }
        public bool OutputClient { get => Data.OutputClient; set => Data.OutputClient = value; }
        public bool OutputServer { get => Data.OutputServer; set => Data.OutputServer = value; }
        public PlatformOutputInfo ClientOutputInfo = new(data.ClientOutputInfo);
        public PlatformOutputInfo ServerOutputInfo = new(data.ServerOutputInfo);
        public LanguageInfo Language = new(data.Language);
        
        public string PreHanleProgramFile { get => ExcelUtil.ParsePath(Data.PreHanleProgramFile); set => Data.PreHanleProgramFile = value; }
        public string PreHanleProgramArgument {get => Data.PreHanleProgramArgument; set => Data.PreHanleProgramArgument = value; }
        public string AftHanleProgramFile { get => ExcelUtil.ParsePath(Data.AftHanleProgramFile); set => Data.AftHanleProgramFile = value; }
        public string AftHanleProgramArgument { get => Data.AftHanleProgramArgument; set => Data.AftHanleProgramArgument = value; }

        public int LastSelectedActionIdx { get => Data.LastSelectedActionIdx; set => Data.LastSelectedActionIdx = value; }
    }
}
