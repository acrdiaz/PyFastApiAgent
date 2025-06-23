using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using dReveal.Common;
using Microsoft.Web.WebView2.WinForms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace dRevealAI
{
    public class EmailListDialog : Form
    {

        #region Fields and properties

        private readonly List<EmailWithSummary> _emails;
        private WebView2 webView;

        public Dictionary<string, Dictionary<string, string>> PromptGroups { get; set; }

        #endregion Fields and properties

        #region Constructor

        public EmailListDialog(List<EmailWithSummary> emailsWithSummaries, Dictionary<string, Dictionary<string, string>> promptGroups)
        {
            _emails = emailsWithSummaries;
            PromptGroups = promptGroups;

            // Must be first
            InitializeComponent();

            // Hook up the Load event
            Load += EmailListDialog_Load;
        }

        #endregion Constructor

        #region HTML handler

        private string GenerateEmailHtml(List<EmailWithSummary> emails)
        {
            var sb = new StringBuilder();
            sb.Append(@"<!DOCTYPE html>
                        <html>
                        <head>
                            <style>
                                body { font-family: 'Segoe UI', sans-serif; margin: 20px; }
                                .email { 
                                    padding: 15px; 
                                    margin-bottom: 15px; 
                                    background: #f8f8f8;
                                    border-left: 4px solid #2b579a;
                                    border-radius: 4px;
                                }
                                .unread { border-left-color: #e74c3c; }
                                .btn { 
                                    padding: 5px 12px;
                                    margin-right: 8px;
                                    border: none;
                                    border-radius: 3px;
                                    cursor: pointer;
                                }
                                .btn-open { background: #2b579a; color: white; }
                                .btn-reply { background: #27ae60; color: white; }
                            </style>
                        </head>
                        <body>");

            foreach (var mail in emails)
            {
                sb.Append($@"
                        <div class='email {(mail.Mail.UnRead ? "unread" : "")}'>
                            <div><strong>📅</strong> {mail.Mail.ReceivedTime:MMM d, yyyy h:mm tt}</div>
                            <div><strong>📩</strong> From: {mail.Mail.SenderName}</div>
                            <div><strong>🔖</strong> Subject: {mail.Mail.Subject}</div>
                            <div><strong>{GetSummaryEmoji(mail.Summary)}</strong> Summary: {mail.Summary}</div>
                            <div style='margin-top:10px;'>
                                <button class='btn btn-open' onclick='window.chrome.webview.postMessage(`open:{mail.Mail.EntryID}`)'>
                                    Open
                                </button>
                                <button class='btn btn-reply' onclick='window.chrome.webview.postMessage(`reply:{mail.Mail.EntryID}`)'>
                                    Reply
                                </button>
                            </div>
                        </div>");
            }

            sb.Append("</body></html>");
            return sb.ToString();
        }

        private void SetupWebViewHandlers()
        {
            try
            {
                webView.CoreWebView2.WebMessageReceived += (sender, e) =>
                {
                    try
                    {
                        var message = e.WebMessageAsJson.Trim('"');
                        var parts = message.Split(':');
                        if (parts.Length == 2)
                        {
                            var action = parts[0];
                            var entryId = parts[1];

                            this.Invoke((MethodInvoker)delegate
                            {
                                switch (action)
                                {
                                    case "open":
                                        OpenEmail(entryId);
                                        break;
                                    case "reply":
                                        ReplyToEmail(entryId);
                                        break;
                                }
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            MessageBox.Show($"Error handling message: {ex.Message}");
                        });
                    }
                };
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error setting up handlers: {ex.Message}");
            }
        }

        #endregion HTML handler

        #region Helpers

        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.ClientSize = new System.Drawing.Size(700, 650);
            this.Name = "EmailListDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "VIP Emails";

            this.ResumeLayout(false);
        }

        private async void EmailListDialog_Load(object sender, EventArgs e)
        {
            try
            {
                webView = new WebView2
                {
                    Dock = DockStyle.Fill,
                    CreationProperties = new CoreWebView2CreationProperties
                    {
                        UserDataFolder = Path.Combine(Path.GetTempPath(), "WebView2Cache")
                    }
                };
                this.Controls.Add(webView);

                await webView.EnsureCoreWebView2Async();

                webView.CoreWebView2.NavigateToString(GenerateEmailHtml(_emails));
                SetupWebViewHandlers();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to initialize WebView2: {ex.Message}");
                this.Close();
            }
        }

        private async Task ReplyToEmail(string entryId)
        {
            Outlook.Application outlookApp = null;
            Outlook.MailItem originalMail = null;
            Outlook.MailItem replyMail = null;

            try
            {
                outlookApp = new Outlook.Application();
                var ns = outlookApp.GetNamespace("MAPI");
                originalMail = ns.GetItemFromID(entryId) as Outlook.MailItem;

                string prompt = string.Format(PromptGroups["DateFilter"]["suggest_formal_reply"], originalMail.Body);
                string draft = await ProcessWithAI(prompt);

                replyMail = originalMail.Reply();

                // Wrap AI response in basic HTML for consistency
                string htmlResponse = $@"
                <div style='font-family:Segoe UI; font-size:10pt; margin-bottom:12px;'>
                    {draft}
                </div>
                <hr style='border:1px solid #ccc;' />";

                string newHtmlBody = htmlResponse + replyMail.HTMLBody;
                replyMail.HTMLBody = newHtmlBody;

                // Show the reply window
                replyMail.Display(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating reply: {ex.Message}");
            }
            finally
            {
                if (replyMail != null) Marshal.ReleaseComObject(replyMail);
                if (originalMail != null) Marshal.ReleaseComObject(originalMail);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }
        }

        private void OpenEmail(string entryId)
        {
            Outlook.Application outlookApp = null;
            Outlook.MailItem mail = null;

            try
            {
                outlookApp = new Outlook.Application();
                var ns = outlookApp.GetNamespace("MAPI");
                mail = ns.GetItemFromID(entryId) as Outlook.MailItem;
                mail.Display(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening email: {ex.Message}");
            }
            finally
            {
                if (mail != null) Marshal.ReleaseComObject(mail);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }
        }

        string GetSummaryEmoji(string summary)
        {
            if (summary.Contains("!")) return "❗";
            if (summary.Contains("action")) return "🎯";
            if (summary.Length < 100) return "💡";
            return "📝";
        }

        #endregion Helpers


        #region Helper Methods

        private readonly AIServiceProvider _aiService = new AIServiceProvider();

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

        #endregion Helper Methods
    }
}
