using dReveal.Common;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace dRevealAI
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        //private readonly AIServiceProvider _aiServiceProvider;
        private readonly AIServiceProvider _aiService = new AIServiceProvider();


        public MyRibbon()
        {
            //_aiServiceProvider = new AIServiceProvider();
        }

        #region Ribbon Handlers
        public void Button_Click(Office.IRibbonControl control)
        {
            var mailItem = GetSelectedMailItem();
            if (mailItem == null) return;

            try
            {
                switch (control.Id)
                {
                    case "btnSummarize":
                        SummarizeEmail(mailItem);
                        break;
                    case "btnSuggestReply":
                        SuggestReply(mailItem);
                        break;
                    case "btnDraftEmail":
                        DraftResponse(mailItem);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in {control.Id.Replace("btn", "")}: {ex.Message}",
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Email Processing
        private async void SummarizeEmail(Outlook.MailItem mailItem)
        {
            string prompt = $"Summarize this email in 3 bullet points:\n\n{mailItem.Body}";
            string summary = await ProcessWithAI(prompt);
            ShowResult("Email Summary", summary);
        }

        private async void SuggestReply(Outlook.MailItem mailItem)
        {
            //string prompt = $"Suggest 3 professional responses to this email:\n\n{mailItem.Body}";
            string prompt =
$"Suggest 3 email replies formatted EXACTLY as:\n" +
$"1. Formal: [Professional tone, full sentences]\n" +
$"2. Neutral: [Concise but polite]\n" +
$"3. Friendly: [Casual with emojis if appropriate]\n\n" +
$"Original email:\n\n{mailItem.Body}";

            string suggestions = await ProcessWithAI(prompt);
            ShowResult("Suggested Replies", suggestions);
        }

        private async void DraftResponse(Outlook.MailItem mailItem)
        {
            string prompt = $"Draft a professional response to this email:\n\n{mailItem.Body}";
            string draft = await ProcessWithAI(prompt);

            Outlook.MailItem newMail = Globals.ThisAddIn.Application
                .CreateItem(Outlook.OlItemType.olMailItem);
            newMail.Subject = "Re: " + mailItem.Subject;
            newMail.Body = draft;
            newMail.Display();
        }
        #endregion

        #region Helper Methods
        private async Task<string> ProcessWithAI(string prompt)
        {
            using (var progressForm = new Form { Text = "Processing...", Width = 300, Height = 100 })
            {
                progressForm.Show();
                progressForm.Refresh();
                var result = await _aiService.GetDefaultService().AnalyzeContentAsync(prompt);

                progressForm.Close();
                return result;
            }
        }

        private Outlook.MailItem GetSelectedMailItem()
        {
            try
            {
                return Globals.ThisAddIn.Application.ActiveInspector()?.CurrentItem as Outlook.MailItem;
            }
            catch
            {
                MessageBox.Show("Please open an email first", "Selection Required",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
        }

        private void ShowResult(string title, string content)
        {
            using (var form = new Form
            {
                Text = title,
                Width = 700,
                Height = 500,
                StartPosition = FormStartPosition.CenterScreen,
                Font = new Font("Segoe UI", 10),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false
            })
            {
                var textBox = new RichTextBox
                {
                    Text = FormatAiResponse(content),
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.None,
                    Margin = new Padding(10),
                    ScrollBars = RichTextBoxScrollBars.Vertical
                };

                var btnCopy = new Button
                {
                    Text = "Copy to Clipboard",
                    Dock = DockStyle.Bottom,
                    Height = 40,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    BackColor = Color.LightGray
                };
                btnCopy.Click += (s, e) => Clipboard.SetText(textBox.Text);

                form.Controls.Add(btnCopy);
                form.Controls.Add(textBox);
                form.ShowDialog();
            }
        }

        private string FormatAiResponse(string rawText)
        {
            // Format AI response with colors and bullets
            return rawText
                .Replace("1. Formal:", "● [Formal] ")
                .Replace("2. Neutral:", "● [Neutral] ")
                .Replace("3. Friendly:", "● [Friendly] ")
                .Replace("\n", "\n    "); // Indent replies
        }
        #endregion


        #region Images resource

        private Bitmap LoadImageFromResource(string resourcePath)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourcePath))
            {
                return stream != null ? new Bitmap(stream) : null;
            }
        }
        
        #endregion Image resource

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("dRevealAI.MyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
