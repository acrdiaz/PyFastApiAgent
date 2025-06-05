using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dReveal.Common
{
    public interface IAIService
    {
        string AnalyzeContent(string input);
        Task<string> AnalyzeContentAsync(string input);
    }
}
