using System;

namespace Grid.Constants
{
    internal static class GridConstants
    {
        public const string ExtensionVersion = "0.1.0";
        public const string ServerName = "grid";
        public const int DefaultMcpPort = 39041;
        public const bool DefaultMcpEnabled = true;
        public const string DefaultProviderBaseUrl = "https://api.openai.com/v1/";
        public const string DefaultProviderModel = "gpt-4.1-mini";
        public const string DefaultSystemPrompt = "You are an Office automation assistant inside Excel. Use the available tools when they are helpful, prefer precise operations over guesswork, and explain the result clearly after using tools.";
        public const int DefaultWordTextLimit = 4000;
        public const int DefaultPowerPointPreviewLimit = 500;
        public const int MaxAgentToolRounds = 8;
        public static readonly TimeSpan ProviderTimeout = TimeSpan.FromSeconds(100);
    }
}
