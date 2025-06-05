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
    public enum DateRange { Today, Yesterday, ThisWeek }

    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private string _selectedDateRange = "today"; // Default


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
                    case "btnListToday":
                        ListTodaysEmails();
                        break;
                    //case "btnListEmails":
                    //    ListEmails_DateRange();
                    //    break;
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
                Width = 750,
                Height = 550,
                StartPosition = FormStartPosition.CenterScreen,
                Font = new Font("Segoe UI", 10),
                //FormBorderStyle = FormBorderStyle.FixedDialog,
                //MaximizeBox = false
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
                    BackColor = Color.LightGray,
                    Enabled = false,
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

        private void ListTodaysEmails()
        {
            Outlook.MAPIFolder inbox = null;
            try
            {
                Outlook.Application outlookApp = Globals.ThisAddIn.Application;
                inbox = outlookApp.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox);

                DateTime today = DateTime.Today;
                var todayEmails = inbox.Items
                    .OfType<Outlook.MailItem>()
                    .Where(mail => mail.ReceivedTime.Date == today)
                    .OrderByDescending(mail => mail.ReceivedTime)
                    .Take(50) // Limit to 50 most recent
                    .ToList();

                if (!todayEmails.Any())
                {
                    MessageBox.Show("No emails found for today.");
                    return;
                }

                // Build formatted list
                var sb = new StringBuilder();
                sb.AppendLine($"📅 Emails Received Today ({today:d})");
                sb.AppendLine("──────────────────────────────");

                foreach (var mail in todayEmails)
                {
                    sb.AppendLine($"• {mail.ReceivedTime:t} - {mail.SenderName}");
                    sb.AppendLine($"  Subject: {mail.Subject}");
                    sb.AppendLine();
                }

                // Display in results window
                ShowResult("Today's Emails", sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error listing emails: {ex.Message}");
            }
            finally
            {
                if (inbox != null)
                {
                    Marshal.ReleaseComObject(inbox);
                }
            }
        }

        // COMBOBOX HANDLERS
        public string GetSelectedDateRange(Office.IRibbonControl control)
        {
            return _selectedDateRange;
        }

        public void OnDateRangeChanged(Office.IRibbonControl control, string selectedId)
        {
            _selectedDateRange = selectedId;
        }

        //ListEmails_DateRange
        //public void ListEmails_DateRange()
        public void ListEmails_Click(Office.IRibbonControl control)
        {
            DateRange dateRange;

            // Explicit switch statement (more reliable in VSTO)
            switch (_selectedDateRange)
            {
                case "Yesterday":
                    dateRange = DateRange.Yesterday;
                    break;
                case "This Week":
                    dateRange = DateRange.ThisWeek;
                    break;
                default: // "today" or any unexpected value
                    dateRange = DateRange.Today;
                    break;
            }

            ListEmailsByDateRange(dateRange);
        }

        private void ListEmailsByDateRange(DateRange range)
        {
            //Outlook.MAPIFolder folder = null;
            //Outlook.Items items = null;
            Outlook.Application outlook = null;
            Outlook.MAPIFolder inbox = null;

            try
            {
                outlook = Globals.ThisAddIn.Application;
                inbox = outlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                
                var (startDate, endDate) = GetDateRange(range);
                string filter = CreateDateFilter(startDate, endDate);

                var emails = inbox.Items.Restrict(filter)
                    .OfType<Outlook.MailItem>()
                    .OrderByDescending(m => m.ReceivedTime)
                    .Take(100)
                    .ToList();

                ShowEmailList(emails, range);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error listing emails: {ex.Message}");
            }
            finally
            {
                //if (inbox != null) Marshal.ReleaseComObject(inbox);
                //if (outlook != null) Marshal.ReleaseComObject(outlook);
            }
        }

        private (DateTime Start, DateTime End) GetDateRange(DateRange range)
        {
            DateTime end = DateTime.Today.AddDays(1).AddSeconds(-1);
            
            switch (range)
            {
                case DateRange.Yesterday:
                    return (end.AddDays(-1).Date, end.AddDays(-1));
                case DateRange.ThisWeek:
                    return (end.AddDays(-(int)end.DayOfWeek).Date, end);
                default: // Today
                    return (end.Date, end);
            }
        }

        private string CreateDateFilter(DateTime start, DateTime end)
        {
            return $"[ReceivedTime] >= '{start:MM/dd/yyyy HH:mm}' AND " +
                   $"[ReceivedTime] <= '{end:MM/dd/yyyy HH:mm}'";
        }

        private void ShowEmailList(List<Outlook.MailItem> emails, DateRange range)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"📅 {range.ToString()} Emails ({emails.Count})");
            sb.AppendLine("──────────────────────────────");

            foreach (var mail in emails)
            {
                sb.AppendLine($"• {mail.ReceivedTime:MMM d h:mm tt} - {mail.SenderName}");
                sb.AppendLine($"  Subject: {mail.Subject}");
                sb.AppendLine($"  {(mail.UnRead ? "🆕 UNREAD" : "✓ Read")}");
                sb.AppendLine();
            }

            ShowResult("Email List", sb.ToString());
        }

        #endregion Helper Methods


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
