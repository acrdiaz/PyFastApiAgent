using System;

namespace dReveal.Common
{
    public class AIServiceProvider
    {
        public IAIService GetDefaultService()
        {
            if (string.IsNullOrEmpty(AISettings.DR_APIKEY))
            {
                throw new InvalidOperationException("API key is not configured");
            }

            switch (AISettings.DR_LLM_DEFAULT)
            {
                case AISettings.DR_LLM_GPT:
                    return new OpenAIService(AISettings.DR_APIKEY);
                case AISettings.DR_LLM_GEMINI:
                    return new GeminiService(AISettings.DR_APIKEY);
                default:
                    throw new InvalidOperationException("Unsupported AI model");
            }
        }
    }
}