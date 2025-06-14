﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dReveal.Common
{
    
    public static class AISettings
    {
        public static string DR_APIKEY = "";
        public const string DR_LLM_GPT = "gpt-4o-mini";
        public const string DR_LLM_GEMINI = "gemini-2.0-flash";
        public static string DR_LLM_DEFAULT = DR_LLM_GEMINI; // AA1 this should be in a .config file

        static AISettings()
        {
            ResolveApiKey();
        }

        private static void ResolveApiKey()
        {
            try
            {
                switch (DR_LLM_DEFAULT)
                {
                    case DR_LLM_GEMINI:
                        DR_APIKEY = GetEnvironmentVariable("GEMINI_API_KEY");
                        break;
                    case DR_LLM_GPT:
                        DR_APIKEY = GetEnvironmentVariable("OPENAI_API_KEY");
                        break;
                    default:
                        DR_APIKEY = string.Empty;
                        break;
                }
            }
            catch (Exception ex)
            {
                DR_APIKEY = string.Empty;
                System.Diagnostics.Debug.WriteLine($"Failed to get API key: {ex.Message}");
            }
        }

        private static string GetEnvironmentVariable(string name)
        {
            string value = Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.User) ??
                           Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Machine);

            if (string.IsNullOrEmpty(value))
            {
                throw new InvalidOperationException($"Environment variable '{name}' not found.");
            }

            return value;
        }
    }
}
