using dReveal.Common;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Text.RegularExpressions;
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

        #region Fields and Properties

        private Office.IRibbonUI ribbon;
        private string SelectedFilterDateRange { get; set; } = "today"; // Default
        private string SelectedVIPDateRange { get; set; } = "today"; // Default
        private string SelectedVIP { get; set; } = string.Empty;

        public enum DateRange { Today, Yesterday, ThisWeek, PreviousSevenDays }

        //private readonly AIServiceProvider _aiServiceProvider;
        private readonly AIServiceProvider _aiService = new AIServiceProvider();

        private List<string> _vipContacts = new List<string>
        {
            "john.doe@company.com",
            "ceo@company.com",
            "important.client@example.com",
            "kabularach@info-arch.com",
            "rcoronado@info-arch.com",
            "cdiaz@info-arch.com",
        };

        // ComboBox callbacks
        public int GetVIPContactCount(Office.IRibbonControl control) => _vipContacts.Count;
        public string GetVIPContactLabel(Office.IRibbonControl control, int index) => _vipContacts[index];
        public string GetSelectedVIP(Office.IRibbonControl control) => SelectedVIP;
        public void OnVIPContactChanged(Office.IRibbonControl control, string selectedId)
            => SelectedVIP = selectedId;

        #endregion region Fields and Properties

        #region Constructor

        public MyRibbon()
        {
            //_aiServiceProvider = new AIServiceProvider();
        }

        #endregion Constructor

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
                    //case "btnListToday":
                    //    ListTodaysEmails();
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

        //private async void ListTodaysEmails()
        //{
        //    Outlook.MAPIFolder inbox = null;
        //    try
        //    {
        //        Outlook.Application outlookApp = Globals.ThisAddIn.Application;
        //        inbox = outlookApp.Session.GetDefaultFolder(
        //            Outlook.OlDefaultFolders.olFolderInbox);

        //        DateTime today = DateTime.Today;
        //        var todayEmails = inbox.Items
        //            .OfType<Outlook.MailItem>()
        //            .Where(mail => mail.ReceivedTime.Date == today)
        //            .OrderByDescending(mail => mail.ReceivedTime)
        //            .Take(50) // Limit to 50 most recent
        //            .ToList();

        //        if (!todayEmails.Any())
        //        {
        //            MessageBox.Show("No emails found for today.");
        //            return;
        //        }

        //        // Build formatted list
        //        var sb = new StringBuilder();
        //        sb.AppendLine($"📅 Emails Received Today ({today:d})");
        //        sb.AppendLine("──────────────────────────────");

        //        string summary = string.Empty;
        //        string prompt = string.Empty;
        //        foreach (var mail in todayEmails)
        //        {
        //            sb.AppendLine($"• {mail.ReceivedTime:t} - {mail.SenderName}");
        //            sb.AppendLine($"  Subject: {mail.Subject}");
        //            prompt = $"Summarize this email in 3 lines:\n\n{mail.Body}";
        //            summary = await ProcessWithAI(prompt);
        //            sb.AppendLine($"  Summary: {summary}");
        //            sb.AppendLine();
        //        }

        //        // Display in results window
        //        ShowResult("Today's Emails", sb.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Error listing emails: {ex.Message}");
        //    }
        //    finally
        //    {
        //        if (inbox != null)
        //        {
        //            Marshal.ReleaseComObject(inbox);
        //        }
        //    }
        //}

        #endregion Email Processing

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
                Font = new System.Drawing.Font("Segoe UI", 10),
                //FormBorderStyle = FormBorderStyle.FixedDialog,
                //MaximizeBox = false
            })
            {
                var textBox = new RichTextBox
                {
                    Text = FormatAiResponse(content),
                    Dock = DockStyle.Fill,
                    ReadOnly = false,
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
                    Font = new System.Drawing.Font("Segoe UI", 9, FontStyle.Bold),
                    BackColor = Color.LightGray,
                    Enabled = false,
                    Visible = false,
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

        // COMBOBOX HANDLERS
        public string GetSelectedDateRange(Office.IRibbonControl control)
        {
            return SelectedFilterDateRange;
        }

        public void OnFilterDateRangeChanged(Office.IRibbonControl control, string selectedId)
        {
            SelectedFilterDateRange = selectedId;
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
                case DateRange.PreviousSevenDays:
                    return (end.AddDays(-7).Date, end.AddDays(-1)); // From 7 days ago to yesterday
                default: // Today
                    return (end.Date, end);
            }
        }

        private string CreateDateFilter(DateTime start, DateTime end, string additionalFilter = null)
        {
            return $"[ReceivedTime] >= '{start:MM/dd/yyyy HH:mm}' AND " +
                   $"[ReceivedTime] <= '{end:MM/dd/yyyy HH:mm}'";
        }

        private string CreateVipDateFilter(DateTime start, DateTime end, string emailAddress)
        {
            //string filter = $"[SenderEmailAddress] = '{emailAddress}' AND " +
            //               $"[ReceivedTime] >= '{DateTime.Today.AddDays(-7):MM/dd/yyyy}'";

            //return $"[SenderEmailAddress] = '{emailAddress}' AND " + CreateDateFilter(start, end);
            //return $"[SenderEmailAddress] = '{emailAddress}'";
            //return $"[SenderEmailAddress] like '{emailAddress}'";
            return $"[SenderEmailAddress] = '{emailAddress}'";
        }

        //private async Task ShowEmailList(List<Outlook.MailItem> emails, DateRange range)
        //{
        //    var sb = new StringBuilder();
        //    sb.AppendLine($"📅 {range.ToString()} Emails ({emails.Count})");
        //    sb.AppendLine("──────────────────────────────");

        //    string prompt = string.Empty;
        //    string summary = string.Empty;
        //    foreach (var mail in emails)
        //    {
        //        // Sanitize the body content
        //        string cleanBody = SanitizeEmailBody(mail.Body);

        //        sb.AppendLine($"• {mail.ReceivedTime:MMM d h:mm tt} - {mail.SenderName}");
        //        sb.AppendLine($"  Subject: {mail.Subject}");
        //        sb.AppendLine($"  {(mail.UnRead ? "🆕 UNREAD" : "✓ Read")}");

        //        if (!string.IsNullOrEmpty(cleanBody))
        //        {
        //            sb.AppendLine($"  Preview: {Truncate(cleanBody, 200)}"); // Show first 200 chars
        //        }

        //        //prompt = $"Summarize this email in 3 lines:\n\n{mail.Body}";
        //        //summary = await ProcessWithAI(prompt);
        //        //sb.AppendLine($"  Summary: {summary}");
        //        //sb.AppendLine($"  Message: {mail.Body}");
        //        //sb.AppendLine($"  Message: {mail.}"); // AA1 is it possible to remove hash?
        //        //mail.BodyFormat
        //        //mail.HTMLBody
        //        //mail.RTFBody

        //        sb.AppendLine();
        //    }

        //    ShowResult("Email List", sb.ToString());
        //}

        private string SanitizeEmailBody(string body)
        {
            if (string.IsNullOrEmpty(body))
                return string.Empty;

            // 1. Remove long hex hashes (e.g., 32+ character strings)
            body = Regex.Replace(body, @"\b[0-9a-fA-F]{32,}\b", "[HASH]");

            // 2. Remove base64-looking strings
            body = Regex.Replace(body, @"\b[A-Za-z0-9+/=]{40,}\b", "[BASE64]");

            // 3. Remove HTML tags if present
            body = Regex.Replace(body, "<[^>]*>", string.Empty);

            // 4. Normalize whitespace
            body = Regex.Replace(body, @"\s+", " ").Trim();

            return body;
        }

        private string Truncate(string value, int maxLength)
        {
            return value.Length <= maxLength ?
                value :
                value.Substring(0, maxLength) + "...";
        }

        #endregion Helper Methods

        #region Date Filter

        public async void FilterListEmails_Click(Office.IRibbonControl control)
        {
            DateRange dateRange;
            switch (SelectedFilterDateRange)
            {
                case "Previous Seven Days":
                    dateRange = DateRange.PreviousSevenDays;
                    break;
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

            await ListEmailsByDateRange(dateRange);
        }

        private async Task ListEmailsByDateRange(DateRange range)
        {
            Outlook.Application outlook = null;
            Outlook.MAPIFolder inbox = null;

            try
            {
                outlook = Globals.ThisAddIn.Application;
                inbox = outlook.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox);

                var (startDate, endDate) = GetDateRange(range);
                string filter = CreateDateFilter(startDate, endDate);

                var emails = inbox.Items.Restrict(filter)
                    .OfType<Outlook.MailItem>()
                    .OrderByDescending(m => m.ReceivedTime)
                    .Take(100)
                    .ToList();

                await ShowEmailList(emails, range);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error listing emails: {ex.Message}");
            }
            finally
            {
            }
        }

        private async Task ShowEmailList(List<Outlook.MailItem> emails, DateRange range)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"📅 {range} Emails ({emails.Count})");
            sb.AppendLine("──────────────────────────────");

            // Process emails in parallel with throttling
            var options = new ParallelOptions { MaxDegreeOfParallelism = 3 };
            var emailTasks = emails.Select(async mail =>
            {
                try
                {
                    var cleanBody = SanitizeEmailBody(mail.Body);
                    if (string.IsNullOrWhiteSpace(cleanBody))
                        return (mail, null);

                    // Get AI summary
                    string prompt = $"Summarize this email in 1-2 sentences. Focus on actions needed and key points:\n\n{cleanBody}";
                    //string summary = await _aiService.GetDefaultService().AnalyzeContentAsync(prompt);
                    string summary = await ProcessWithAI(prompt);

                    return (mail, summary);
                }
                catch
                {
                    return (mail, null); // Fallback if summary fails
                }
            });

            var summarizedEmails = await Task.WhenAll(emailTasks);

            foreach (var (mail, summary) in summarizedEmails)
            {
                sb.AppendLine($"• {mail.ReceivedTime:MMM d h:mm tt} - {mail.SenderName}");
                sb.AppendLine($"  Subject: {mail.Subject}");
                sb.AppendLine($"  {(mail.UnRead ? "🆕 UNREAD" : "✓ Read")}");

                if (!string.IsNullOrEmpty(summary))
                {
                    sb.AppendLine($"  Summary: {summary}");
                }
                else
                {
                    sb.AppendLine($"  Preview: {Truncate(SanitizeEmailBody(mail.Body), 100)}");
                }

                sb.AppendLine();
            }

            ShowResult("Email List", sb.ToString());
        }

        #endregion Date Filter

        #region VIP emails

        public async void CheckVIPEmails_Click(Office.IRibbonControl control)
        {
            if (string.IsNullOrEmpty(SelectedVIP))
            {
                MessageBox.Show("Please select a VIP contact first");
                return;
            }

            DateRange dateRange;
            switch (SelectedVIPDateRange)
            {
                case "Previous Seven Days":
                    dateRange = DateRange.PreviousSevenDays;
                    break;
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

            try
            {
                var emails = await GetEmailsFromVIP(SelectedVIP, dateRange);
                await ShowVIPEmailList(emails, SelectedVIP);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking VIP emails: {ex.Message}");
            }
        }

        // OnVIPDateRangeChanged
        public void OnVIPDateRangeChanged(Office.IRibbonControl control, string selectedId)
        {
            SelectedVIPDateRange = selectedId;
        }

        private async Task<List<Outlook.MailItem>> GetEmailsFromVIP(
            string emailAddress, 
            DateRange range)
        {
            return await Task.Run(() =>
            {
                Outlook.Application outlook = Globals.ThisAddIn.Application;
                Outlook.MAPIFolder inbox = outlook.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox);

                //--
                //foreach (var item in inbox.Items.OfType<Outlook.MailItem>().Take(50))
                //{
                //    var mail = item as Outlook.MailItem;
                //    if (mail != null)
                //    {
                //        Trace.WriteLine($"From: {mail.SenderEmailAddress} | Subject: {mail.Subject}");
                //    }
                //}

                // DateRange range = DateRange.ThisWeek;
                var (startDate, endDate) = GetDateRange(range);
                //string filter = CreateVipDateFilter(startDate, endDate, emailAddress);
                string filter = CreateDateFilter(startDate, endDate);

                var emails = inbox.Items.Restrict(filter)
                    .OfType<Outlook.MailItem>()
                    .OrderByDescending(m => m.ReceivedTime)
                    .Take(50) // Limit results
                    .ToList();

                Marshal.ReleaseComObject(inbox); // AA1 is this necessary?
                return emails;
            });
        }

        private async Task ShowVIPEmailList(List<Outlook.MailItem> emails, string vipEmail)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"⭐ VIP Emails From: {vipEmail}");
            sb.AppendLine($"📅 Last ? Days ({emails.Count} emails)");
            sb.AppendLine("──────────────────────────────");

            foreach (var mail in emails)
            {
                //string summary = await GetEmailSummary(mail);

                sb.AppendLine($"• {mail.ReceivedTime:MMM d h:mm tt}");
                sb.AppendLine($"  Sender: {mail.Sender}");
                sb.AppendLine($"  SenderEmailAddress: {mail.SenderEmailAddress}");
                sb.AppendLine($"  SenderName: {mail.SenderName}");
                sb.AppendLine($"  Subject: {mail.Subject}");
                sb.AppendLine($"  Status: {(mail.UnRead ? "UNREAD" : "Read")}");
                //sb.AppendLine($"  Summary: {summary}"); AA1 enable this
                sb.AppendLine($"  Importance: {mail.Importance.ToString().ToUpper()}");
                sb.AppendLine();
            }

            ShowResult($"VIP Emails from {vipEmail}", sb.ToString());
        }

        private async Task<string> GetEmailSummary(Outlook.MailItem mail)
        {
            var cleanBody = SanitizeEmailBody(mail.Body);
            string prompt = $"Concise summary focusing on actions and deadlines:\n{cleanBody}";
            return await _aiService.GetDefaultService().AnalyzeContentAsync(prompt);
        }

        #endregion VIP emails


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
            //this.ribbon.Invalidate(); // Forces the ribbon to refresh
        }

        //// AA1 this does not work
        //// Callback to get the selected item ID for cmbFilterDateRange
        //public string GetSelectedItemID(Office.IRibbonControl control)
        //{
        //    if (control.Id == "cmbFilterDateRange")
        //    {
        //        return "today"; // Default value
        //    }
        //    return null;
        //}

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
