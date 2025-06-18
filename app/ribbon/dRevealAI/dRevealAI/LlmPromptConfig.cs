using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace dRevealAI
{
    public class LlmPromptConfig
    {
        public Dictionary<string, Dictionary<string, string>> PromptGroups { get; set; }

        public string LlmPromptFile { get; set; } = "LlmPrompts.json";

        public LlmPromptConfig LoadPrompts()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(baseDir, LlmPromptFile);

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Prompt file not found at expected location: {filePath}");

            string json = File.ReadAllText(filePath);
            return JsonConvert.DeserializeObject<LlmPromptConfig>(json);
        }
    }
}