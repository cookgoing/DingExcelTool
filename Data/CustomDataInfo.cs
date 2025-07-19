namespace DingExcelTool.Data
{
    internal class CustomDataInfo
    {
        public class PlatformOutputInfo
        {
            public ScriptTypeEn ScriptType;
            public string ProtoMetaOutputDir;
            public string ProtoScriptOutputDir;
            public string ProtoDataOutputDir;
            public string ExcelScriptOutputDir;
            public string ErrorCodeFrameDir;
            public string ErrorCodeBusinessDir;
        }

        public class LanguageInfo
        {
            public string LanguageExcelDir;
            public string ScriptRootDir;
            public string UIRootDir;
            public string LanguageBackupDir;
            public LanguageType SourceLanguage;
            public string LanguageReplaceMethod;
            public string ImageReplaceMethod;
        }

        public string ExcelInputRootDir;
        public bool OutputClient;
        public bool OutputServer;
        public PlatformOutputInfo ClientOutputInfo;
        public PlatformOutputInfo ServerOutputInfo;
        public LanguageInfo Language;

        public string PreHanleProgramFile;
        public string PreHanleProgramArgument;
        public string AftHanleProgramFile;
        public string AftHanleProgramArgument;

        public int LastSelectedActionIdx;
    }
}
