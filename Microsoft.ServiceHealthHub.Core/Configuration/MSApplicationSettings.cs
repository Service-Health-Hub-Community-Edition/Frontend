namespace Microsoft.ServiceHealthHub.Core
{
    public class MSApplicationBrandingSettings
    {
        public bool enabled = false;
        public string backgroundColor = string.Empty;
        public string logo = string.Empty;
    }

    public class MSApplicationSettings
    {
        public string pirStorageUri = string.Empty;
        public string releaseMessageEndpoint = string.Empty;
        public bool azureTranslatorEnabled = false;
        public string azureTranslatorLocation = string.Empty;
        public string azureTranslatorKey = string.Empty;
        public string languageServiceEndpoint = string.Empty;
        public string languageServiceKey = string.Empty;
        public bool languageServiceEnabled = false;
        public int textSummarizationSentenceCount = 3;
        public MSApplicationBrandingSettings branding = new MSApplicationBrandingSettings();
    }
}