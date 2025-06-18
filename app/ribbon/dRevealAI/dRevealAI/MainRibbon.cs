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

using Microsoft.Web.WebView2.WinForms;

namespace dRevealAI
{

    [ComVisible(true)]
    public class MainRibbon : Office.IRibbonExtensibility
    {

        #region Enums

        public enum DateRange
        {
            Today,
            Yesterday,
            ThisWeek,
            PreviousSevenDays
        }

        #endregion Enums

        #region Fields and Properties

        private Office.IRibbonUI ribbon;

        private string SelectedFilterDateRange { get; set; } = "today"; // Default
        
        private string SelectedVIP { get; set; } = string.Empty;

        private readonly AIServiceProvider _aiService = new AIServiceProvider();

        private List<string> _vipContacts = new List<string>
        {
            "<All>",
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

        public Dictionary<string, Dictionary<string, string>> PromptGroups { get; set; }

        #endregion Fields and Properties

        #region Constructor

        public MainRibbon()
        {
            PromptGroups = new LlmPromptConfig().LoadPrompts().PromptGroups;
        }

        #endregion Constructor

        #region Ribbon Email AI Tools

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

        #endregion Ribbon Email AI Tools

        #region Email Processing

        private async void SummarizeEmail(Outlook.MailItem mailItem)
        {
            string prompt = string.Format(PromptGroups["EmailAITools"]["summarize"], mailItem.Body);

            string summary = await ProcessWithAI(prompt);
            ShowResult("Email Summary", summary);
        }

        private async void SuggestReply(Outlook.MailItem mailItem)
        {
            string prompt = string.Format(PromptGroups["EmailAITools"]["suggest_reply"], mailItem.Body);

            string suggestions = await ProcessWithAI(prompt);
            ShowResult("Suggested Replies", suggestions);
        }

        private async void DraftResponse(Outlook.MailItem mailItem)
        {
            if (mailItem == null)
                return;

            try
            {
                // Generate AI draft based on original body
                string prompt = string.Format(PromptGroups["EmailAITools"]["draft_response"], mailItem.Body);
                string draft = await ProcessWithAI(prompt);

                var reply = mailItem.Reply();

                // Wrap AI response in basic HTML for consistency
                string htmlResponse = $@"
                <div style='font-family:Segoe UI; font-size:10pt; margin-bottom:12px;'>
                    {draft}
                </div>
                <hr style='border:1px solid #ccc;' />";

                string newHtmlBody = htmlResponse + reply.HTMLBody;
                reply.HTMLBody = newHtmlBody;

                // Show the reply window
                reply.Display(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating reply: {ex.Message}");
            }
        }

        // AA1 remove candidate
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

        //private void ShowResult(string title, string content, List<Outlook.MailItem> emails)
        //{
        //    using (var form = new Form
        //    {
        //        Text = title,
        //        Width = 800,
        //        Height = 600,
        //        StartPosition = FormStartPosition.CenterScreen,
        //        Font = new Font("Segoe UI", 10),
        //        FormBorderStyle = FormBorderStyle.FixedDialog,
        //        MaximizeBox = false,
        //        MinimizeBox = false
        //    })
        //    {
        //        var panel = new FlowLayoutPanel
        //        {
        //            Dock = DockStyle.Fill,
        //            AutoScroll = true,
        //            WrapContents = false,
        //            FlowDirection = FlowDirection.TopDown
        //        };

        //        var header = new Label
        //        {
        //            Text = $"⭐ VIP Emails from: {title}\n({emails.Count} emails)",
        //            Font = new Font("Segoe UI", 12, FontStyle.Bold),
        //            AutoSize = true,
        //            Margin = new Padding(10, 10, 0, 10)
        //        };
        //        panel.Controls.Add(header);

        //        foreach (var mail in emails)
        //        {
        //            string entryID = mail.EntryID;
        //            string storeID = mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null;

        //            var itemPanel = new Panel
        //            {
        //                Width = panel.Width - 40,
        //                MinimumSize = new Size(0, 60),
        //                BorderStyle = BorderStyle.FixedSingle,
        //                Padding = new Padding(10),
        //                Cursor = Cursors.Hand
        //            };

        //            var lblSubject = new Label
        //            {
        //                Text = mail.Subject,
        //                Font = new Font("Segoe UI", 9.75f, FontStyle.Bold),
        //                AutoSize = true
        //            };

        //            var lblInfo = new Label
        //            {
        //                Text = $"{mail.ReceivedTime:MMM d h:mm tt} | {(mail.UnRead ? "UNREAD" : "Read")}",
        //                ForeColor = mail.UnRead ? Color.DarkBlue : Color.Gray,
        //                AutoSize = true
        //            };

        //            itemPanel.Controls.Add(lblSubject);
        //            itemPanel.Controls.Add(lblInfo);
        //            lblInfo.Location = new Point(0, lblSubject.Height + 4);

        //            itemPanel.Click += (s, e) => OpenEmail(entryID, storeID);

        //            panel.Controls.Add(itemPanel);
        //        }

        //        var btnCopy = new Button
        //        {
        //            Text = "Copy All to Clipboard",
        //            Width = 180,
        //            Height = 35,
        //            Margin = new Padding(10),
        //            Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        //        };
        //        btnCopy.Click += (s, e) => Clipboard.SetText(content);

        //        form.Controls.Add(panel);
        //        form.Controls.Add(btnCopy);
        //        form.AcceptButton = btnCopy;
        //        form.ShowDialog();
        //    }
        //}

        private void ShowResult(string title, string content, List<Outlook.MailItem> emails)
        {
            using (var form = new Form())
            {
                var rtb = new RichTextBox
                {
                    Dock = DockStyle.Fill,
                    DetectUrls = true
                };

                // Build content with clickable links
                var sb = new StringBuilder();
                for (int i = 0; i < emails.Count; i++)
                {
                    sb.AppendLine($"📧 {emails[i].Subject}");
                    sb.AppendLine($"🔗 Open Email [id:{emails[i].EntryID}]"); // Link pattern
                    sb.AppendLine();
                }

                rtb.Text = sb.ToString();
                rtb.LinkClicked += (s, e) =>
                {
                    var match = Regex.Match(e.LinkText, @"^id:(.*)");
                    if (match.Success) OpenEmailById(match.Groups[1].Value);
                };

                form.Controls.Add(rtb);
                form.ShowDialog();
            }
        }

        private void OpenEmailById(string entryId)
        {
            var mail = (Outlook.MailItem)Globals.ThisAddIn.Application.Session.GetItemFromID(entryId);
            mail.Display(false);
            Marshal.ReleaseComObject(mail);
        }

        private void OpenEmail(string entryID, string storeID)
        {
            if (string.IsNullOrEmpty(entryID) || string.IsNullOrEmpty(storeID))
            {
                MessageBox.Show("Unable to open email: missing EntryID or StoreID.");
                return;
            }

            try
            {
                var app = Globals.ThisAddIn.Application;
                var mailItem = app.Session.GetItemFromID(entryID, storeID) as Outlook.MailItem;
                if (mailItem != null)
                {
                    mailItem.Display(false); // false = open in same thread
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening email: {ex.Message}");
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

        //private string CreateVipDateFilter(DateTime start, DateTime end, string emailAddress)
        //{
        //    //string filter = $"[SenderEmailAddress] = '{emailAddress}' AND " +
        //    //               $"[ReceivedTime] >= '{DateTime.Today.AddDays(-7):MM/dd/yyyy}'";

        //    //return $"[SenderEmailAddress] = '{emailAddress}' AND " + CreateDateFilter(start, end);
        //    //return $"[SenderEmailAddress] = '{emailAddress}'";
        //    //return $"[SenderEmailAddress] like '{emailAddress}'";
        //    return $"[SenderEmailAddress] = '{emailAddress}'";
        //}

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

        #region Ribbon Date Filter

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

        #endregion Ribbon Date Filter

        #region VIP emails

        public async void CheckVIPEmails_Click(Office.IRibbonControl control)
        {
            if (string.IsNullOrEmpty(SelectedVIP))
            {
                MessageBox.Show("Please select a VIP contact first");
                return;
            }

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

            try
            {
                var emails = await GetEmailsFromVIP(SelectedVIP, dateRange);
                await ShowVIPEmailList(emails, SelectedVIP, dateRange);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking VIP emails: {ex.Message}");
            }
        }

        private async Task<List<Outlook.MailItem>> GetEmailsFromVIP(string emailAddress, DateRange range)
        {
            return await Task.Run(() =>
            {
                Outlook.Application outlook = Globals.ThisAddIn.Application;
                Outlook.MAPIFolder inbox = null;
                try
                {
                    inbox = outlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                    var (startDate, endDate) = GetDateRange(range);
                    string filter = CreateDateFilter(startDate, endDate);

                    var filteredEmails = inbox.Items.Restrict(filter)
                        .OfType<Outlook.MailItem>()
                        .OrderByDescending(m => m.ReceivedTime)
                        .Take(50)
                        .ToList();

                    if (emailAddress == "<All>")
                    {
                        return filteredEmails;
                    }

                    var result = new List<Outlook.MailItem>();
                    foreach (var mail in filteredEmails)
                    {
                        string smtpAddress = GetSmtpAddress(mail);

                        if (string.Equals(smtpAddress, emailAddress, StringComparison.OrdinalIgnoreCase))
                        {
                            result.Add(mail);
                        }
                    }

                    return result;
                }
                catch (Exception ex)
                {
                    // AA1 Log or show error message
                    Console.WriteLine($"Error retrieving VIP emails: {ex.Message}");
                    return new List<Outlook.MailItem>();
                }
                finally
                {
                    if (inbox != null)
                    {
                        Marshal.ReleaseComObject(inbox);
                    }
                }
            });
        }

        private string GetSmtpAddress(Outlook.MailItem mailItem)
        {
            if (mailItem == null)
                throw new ArgumentNullException(nameof(mailItem));

            try
            {
                // If it's an Exchange user, resolve the SMTP address
                if (mailItem.SenderEmailType == "EX")
                {
                    var sender = mailItem.Sender;
                    if (sender != null)
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        return sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                    }
                }
                else
                {
                    // Regular SMTP sender
                    return mailItem.SenderEmailAddress;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error resolving SMTP address: {ex.Message}");
            }

            return null;
        }

        private async Task ShowVIPEmailList(List<Outlook.MailItem> emails, string vipEmail, DateRange range)
        {
            if (emails == null || emails.Count == 0)
            {
                MessageBox.Show("No recent emails found from this VIP.");
                return;
            }

            //var sb = new StringBuilder();
            //sb.AppendLine($"⭐ VIP Emails From: {vipEmail}");
            //sb.AppendLine($"📅 {range} Emails ({emails.Count})");
            //sb.AppendLine("──────────────────────────────");

            //foreach (var mail in emails)
            //{
            //    //string summary = await GetEmailSummary(mail); AA1 enable this
            //    string summary = string.Empty;

            //    sb.AppendLine($"• {mail.ReceivedTime:MMM d h:mm tt}");
            //    sb.AppendLine($"  SenderName: {mail.SenderName}");
            //    sb.AppendLine($"  Subject: {mail.Subject}");
            //    sb.AppendLine($"  Status: {(mail.UnRead ? "UNREAD" : "Read")}");
            //    sb.AppendLine($"  Importance: {mail.Importance.ToString().ToUpper()}");
            //    sb.AppendLine($"  Summary: {summary}");
            //    //sb.AppendLine($"  ID: {mail.ConversationID}"); // Implement "conversation view"
            //    //sb.AppendLine($"  ID: {mail.EntryID}"); // Track a specific email
            //    sb.AppendLine();
            //}

            // Pass both the text and the list of emails
            //ShowResult($"VIP Emails from {vipEmail}", sb.ToString(), emails);
            //ShowResult($"VIP Emails from {vipEmail}", sb.ToString());
            //ShowHybridView(emails);
            //ShowEmailListWeb(emails);
            //await ShowEmailListWebView2(emails);


            //var emailsWithSummaries = await Task.WhenAll(emails.Select(async mail =>
            //new {
            //    Mail = mail,
            //    Summary = await GetEmailSummary(mail)
            //}));

            var emailsWithSummaries = await Task.WhenAll(
                emails.Select(async mail => new EmailWithSummary
                {
                    Mail = mail,
                    Summary = await GetEmailSummary(mail)
                })
            );


            await Task.Run(() =>
            {
                var thread = new Thread(() =>
                {
                    Application.EnableVisualStyles();
                    //var form = new EmailListForm(emails);
                    var form = new EmailListDialog(emailsWithSummaries.ToList());
                    //List<(Outlook.MailItem Mail, string Summary)>
                    Application.Run(form);
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
            });

        }

        private async Task ShowEmailListWebView2(List<Outlook.MailItem> emails)
        {
            // Create the form
            var form = new Form
            {
                Text = "VIP Emails",
                Width = 700,
                Height = 650,
                StartPosition = FormStartPosition.CenterScreen
            };

            // Create and configure WebView2
            var webView = new WebView2
            {
                Dock = DockStyle.Fill,
                CreationProperties = new CoreWebView2CreationProperties
                {
                    UserDataFolder = Path.Combine(Path.GetTempPath(), "WebView2Cache")
                }
            };
            form.Controls.Add(webView);

            // Initialize WebView2
            await webView.EnsureCoreWebView2Async();

            // Load HTML content
            webView.CoreWebView2.NavigateToString(GenerateEmailHtml(emails));

            // Handle click events from JavaScript
            webView.CoreWebView2.WebMessageReceived += (sender, e) =>
            {
                if (Guid.TryParse(e.WebMessageAsJson.Trim('"'), out var entryId))
                {
                    OpenEmail(entryId.ToString());
                }
            };

            form.ShowDialog();
        }

        private string GenerateEmailHtml(List<Outlook.MailItem> emails)
        {
            var sb = new StringBuilder();
            sb.Append(@"<!DOCTYPE html>
    <html>
    <head>
        <style>
            body { 
                font-family: 'Segoe UI', sans-serif; 
                margin: 20px;
                line-height: 1.6;
            }
            .email {
                padding: 15px;
                margin-bottom: 15px;
                border-radius: 5px;
                background: #f9f9f9;
                border-left: 4px solid #2b579a;
            }
            .action-btn {
                background: #2b579a;
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 3px;
                cursor: pointer;
                margin-top: 5px;
            }
            .unread {
                border-left-color: #e74c3c;
                background: #fff0f0;
            }
        </style>
    </head>
    <body>");

            foreach (var mail in emails)
            {
                sb.Append($@"
        <div class='email {(mail.UnRead ? "unread" : "")}'>
            <div><strong>📅</strong> {mail.ReceivedTime:MMM d, yyyy h:mm tt}</div>
            <div><strong>📩</strong> From: {mail.SenderName}</div>
            <div><strong>🔖</strong> Subject: {mail.Subject}</div>
            <div><strong>📌</strong> Status: {(mail.UnRead ? "UNREAD" : "Read")}</div>
            <button class='action-btn' onclick='window.chrome.webview.postMessage(`{mail.EntryID}`)'>
                Open Email
            </button>
        </div>");
            }

            sb.Append("</body></html>");
            return sb.ToString();
        }

        private void OpenEmail(string entryId)
        {
            try
            {
                Outlook.MailItem mail = null;
                try
                {
                    mail = Globals.ThisAddIn.Application.Session.GetItemFromID(entryId) as Outlook.MailItem;
                    mail?.Display(false);
                }
                finally
                {
                    if (mail != null) Marshal.ReleaseComObject(mail);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening email:\n{ex.Message}", "Error",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }







        //// Helper methods
        //private void OpenEmail(string entryId)
        //{
        //    try
        //    {
        //        var mail = (Outlook.MailItem)Globals.ThisAddIn.Application.Session.GetItemFromID(entryId);
        //        mail.Display(false);
        //        Marshal.ReleaseComObject(mail);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Failed to open email:\n{ex.Message}");
        //    }
        //}

        private void ReplyToEmail(string entryId)
        {
            try
            {
                var mail = (Outlook.MailItem)Globals.ThisAddIn.Application.Session.GetItemFromID(entryId);
                var reply = mail.Reply();
                reply.Display(false);
                Marshal.ReleaseComObject(reply);
                Marshal.ReleaseComObject(mail);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to reply:\n{ex.Message}");
            }
        }

        private async Task<string> GetEmailSummary(Outlook.MailItem mail)
        {
            var cleanBody = SanitizeEmailBody(mail.Body);
            //Return ONLY this exact HTML format (no other text/comments):
            //string prompt = $"Return ONLY this exact <div> (no other text/comments/wrappers): Concise summary focusing on actions and deadlines:\n{cleanBody}";
            //string prompt = $"Return ONLY the exact original <div> don´t add markdown comments, exclude signature at email footer: Concise summary focusing on actions and deadlines:\n{cleanBody}";
            string prompt = $"Return html <div> snippet. Concise summary in 2 lines focusing on actions and deadlines:\n{cleanBody}";

//            string prompt = $@"
//In html format only the <div> part, exclude the <html>, <head>
//Create two summaries for this email:

//🚀 QUICK ACTION SUMMARY (1 line):
//- Only critical actions/deadlines
//- Format: [emoji] [action] [deadline]

//📌 DETAILED SUMMARY (2 lines):
//- Context + actions + owners
//- Format: • [Point 1]
//         • [Point 2]

//Content:
//{cleanBody}";

//            string prompt = $@"
//Return ONLY the following HTML div (no <html>/<head>/<body>):
//<div>
// <b>🚀 [One-line action]</b><br>
// • [Point 1]<br>
// • [Point 2]
//</div>

//For email:
//{cleanBody}";
            //return await _aiService.GetDefaultService().AnalyzeContentAsync(prompt);
            string result = await _aiService.GetDefaultService().AnalyzeContentAsync(prompt);
            string resultA = RemoveMarkdownCodeBlocks(result);
            return resultA;
        }

        public static string RemoveMarkdownCodeBlocks(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return input;

            // Pattern matches:
            // 1. Optional whitespace
            // 2. ``` followed by optional "html"
            // 3. Optional whitespace
            string pattern = @"(?:\s*```(?:html)?\s*)|(?:\s*```\s*$)";

            return Regex.Replace(input, pattern, "",
                   RegexOptions.Multiline | RegexOptions.IgnoreCase).Trim();
        }

        #endregion VIP emails

        #region Legacv code



        #endregion Legacv code


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
