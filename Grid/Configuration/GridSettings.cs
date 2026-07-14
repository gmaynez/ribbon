using System;
using System.Configuration;
using System.Security.Cryptography;
using System.Text;
using Grid.Constants;

namespace Grid.Configuration
{
    internal sealed class GridSettings : ApplicationSettingsBase
    {
        private static readonly GridSettings DefaultInstance = (GridSettings)Synchronized(new GridSettings());

        public static GridSettings Default
        {
            get { return DefaultInstance; }
        }

        [UserScopedSetting]
        [DefaultSettingValue("True")]
        public bool McpEnabled
        {
            get { return (bool)this[nameof(McpEnabled)]; }
            set { this[nameof(McpEnabled)] = value; }
        }

        [UserScopedSetting]
        [DefaultSettingValue("39041")]
        public int McpPort
        {
            get { return (int)this[nameof(McpPort)]; }
            set { this[nameof(McpPort)] = value; }
        }

        [UserScopedSetting]
        [DefaultSettingValue(GridConstants.DefaultProviderBaseUrl)]
        public string ProviderBaseUrl
        {
            get { return (string)this[nameof(ProviderBaseUrl)]; }
            set { this[nameof(ProviderBaseUrl)] = value; }
        }

        [UserScopedSetting]
        [DefaultSettingValue(GridConstants.DefaultProviderModel)]
        public string ProviderModel
        {
            get { return (string)this[nameof(ProviderModel)]; }
            set { this[nameof(ProviderModel)] = value; }
        }

        [UserScopedSetting]
        [DefaultSettingValue(GridConstants.DefaultSystemPrompt)]
        public string SystemPrompt
        {
            get { return (string)this[nameof(SystemPrompt)]; }
            set { this[nameof(SystemPrompt)] = value; }
        }

        [UserScopedSetting]
        [DefaultSettingValue("")]
        public string ProtectedProviderApiKey
        {
            get { return (string)this[nameof(ProtectedProviderApiKey)]; }
            set { this[nameof(ProtectedProviderApiKey)] = value; }
        }

        public string GetProviderApiKey()
        {
            byte[] unprotectedBytes;
            byte[] protectedBytes;

            if (string.IsNullOrWhiteSpace(ProtectedProviderApiKey))
            {
                return string.Empty;
            }

            try
            {
                protectedBytes = Convert.FromBase64String(ProtectedProviderApiKey);
                unprotectedBytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(unprotectedBytes);
            }
            catch
            {
                return string.Empty;
            }
        }

        public void SetProviderApiKey(string apiKey)
        {
            byte[] plainBytes;
            byte[] protectedBytes;

            if (string.IsNullOrWhiteSpace(apiKey))
            {
                ProtectedProviderApiKey = string.Empty;
                return;
            }

            plainBytes = Encoding.UTF8.GetBytes(apiKey.Trim());
            protectedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
            ProtectedProviderApiKey = Convert.ToBase64String(protectedBytes);
        }

        public void Normalize()
        {
            if (McpPort < 1024 || McpPort > 65535)
            {
                McpPort = GridConstants.DefaultMcpPort;
            }

            if (string.IsNullOrWhiteSpace(ProviderBaseUrl))
            {
                ProviderBaseUrl = GridConstants.DefaultProviderBaseUrl;
            }

            if (string.IsNullOrWhiteSpace(ProviderModel))
            {
                ProviderModel = GridConstants.DefaultProviderModel;
            }

            if (string.IsNullOrWhiteSpace(SystemPrompt))
            {
                SystemPrompt = GridConstants.DefaultSystemPrompt;
            }
        }

        public void SaveNormalized()
        {
            Normalize();
            Save();
        }
    }
}
