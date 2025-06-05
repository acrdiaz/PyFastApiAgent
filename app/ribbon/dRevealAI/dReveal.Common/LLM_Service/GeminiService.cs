using GenerativeAI;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace dReveal.Common
{
    public class GeminiService : IAIService
    {
        private readonly string _apiKey;
        private GenerativeModel _model;

        public GeminiService(string apiKey)
        {
            _apiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));
            InitializeModel();
        }

        private void InitializeModel()
        {
            try
            {
                var client = new GenerativeModel(
                    apiKey: _apiKey,
                    model: AISettings.DR_LLM_GEMINI);

                _model = client;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to initialize Gemini model: {ex.Message}",
                    "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        public string AnalyzeContent(string input)
        {
            return AnalyzeContentAsync(input).GetAwaiter().GetResult();
        }

        public async Task<string> AnalyzeContentAsync(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "Error: Input text is empty";

            try
            {
                var response = await _model.GenerateContentAsync(input);
                return response.Text;
            }
            catch (Exception ex)
            {
                // Handle specific Gemini API errors
                if (ex.Message.Contains("API key not valid"))
                {
                    MessageBox.Show("Invalid Gemini API key. Please check your environment variables.",
                        "API Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (ex.Message.Contains("quota"))
                {
                    MessageBox.Show("API quota exceeded. Please check your Google Cloud quota.",
                        "Quota Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                return $"AI Service Error: {ex.Message}";
            }
        }
    }
}