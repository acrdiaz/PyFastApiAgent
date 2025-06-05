using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace dReveal.Common
{
    public static class FeedbackLogger
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "dRevealAI\\feedback.log");

        public static void Log(string rating, string content)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));

                // Truncate long content and sanitize
                string cleanContent = new string(content.Take(500).ToArray())
                    .Replace("\r", "").Replace("\n", " ");

                string logEntry = $"{DateTime.UtcNow:o}|{rating}|{cleanContent}\n";

                File.AppendAllText(LogPath, logEntry);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Feedback logging failed: {ex.Message}");
            }
        }

        public static void AnalyzeFeedback()
        {
            if (!File.Exists(LogPath)) return;

            var logs = File.ReadAllLines(LogPath)
                .Select(line => line.Split('|'))
                .Where(parts => parts.Length == 3);

            int positive = logs.Count(x => x[1] == "good");
            int negative = logs.Count(x => x[1] == "bad");

            // Optional: Use this data to improve your AI prompts
        }
    }
}
