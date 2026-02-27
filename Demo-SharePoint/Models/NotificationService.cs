using System;
using System.Configuration;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace DEMO_SharePoint.Models
{
    /// <summary>
    /// Sends workflow e-mail notifications via SMTP.
    /// Configure keys in Web.config appSettings:
    ///   Smtp.Host, Smtp.Port, Smtp.From, Smtp.FromName,
    ///   Smtp.Username, Smtp.Password, Smtp.EnableSsl, Smtp.Domain
    /// If Smtp.Host is empty the service silently no-ops.
    /// </summary>
    public class NotificationService
    {
        private readonly string _host;
        private readonly int _port;
        private readonly string _from;
        private readonly string _fromName;
        private readonly string _smtpUser;
        private readonly string _smtpPass;
        private readonly bool _ssl;
        private readonly string _smtpDomain;
        private readonly bool _enabled;

        public NotificationService()
        {
            _host      = ConfigurationManager.AppSettings["Smtp.Host"] ?? "";
            _port      = int.TryParse(ConfigurationManager.AppSettings["Smtp.Port"], out var p) ? p : 25;
            _from      = ConfigurationManager.AppSettings["Smtp.From"] ?? "dms@company.com";
            _fromName  = ConfigurationManager.AppSettings["Smtp.FromName"] ?? "AppKings DMS";
            _smtpUser  = ConfigurationManager.AppSettings["Smtp.Username"] ?? "";
            _smtpPass  = ConfigurationManager.AppSettings["Smtp.Password"] ?? "";
            _ssl       = string.Equals(ConfigurationManager.AppSettings["Smtp.EnableSsl"], "true",
                             StringComparison.OrdinalIgnoreCase);
            _smtpDomain = ConfigurationManager.AppSettings["Smtp.Domain"] ?? "";
            _enabled   = !string.IsNullOrWhiteSpace(_host);
        }

        // Public notification methods

        public void NotifySubmitted(string approverEmail, string itemName,
                                    string submittedBy, int level, string siteBaseUrl, string itemUrl)
        {
            if (!_enabled) return;
            var subject = $"[Action Required] Document submitted for your approval – {itemName}";
            var body = BuildBody(
                title:    "Approval Required",
                iconColor:"#E84F26",
                icon:     "&#128221;",
                lines:    new[]
                {
                    $"<strong>{submittedBy}</strong> has submitted a document for your approval at <strong>Level {level}</strong>.",
                    $"<br/>Document: <strong>{itemName}</strong>",
                    $"<br/>Please log in to the DMS to review and approve or reject."
                },
                actionUrl:  siteBaseUrl + "/Workflow/MyApprovals",
                actionLabel:"Review Document",
                footerNote: $"You are receiving this because you are an approver at Level {level}."
            );
            SendEmail(approverEmail, subject, body);
        }

        public void NotifyApproved(string submitterEmail, string itemName,
                                   string approver, int level, int totalLevels)
        {
            if (!_enabled) return;
            bool isFinal = level == totalLevels;
            var subject = isFinal
                ? $"[Approved] {itemName} - all approval levels complete"
                : $"[Approved] {itemName} - Level {level} approved";
            var body = BuildBody(
                title:    isFinal ? "Workflow Complete" : $"Level {level} Approved",
                iconColor:"#83BB00",
                icon:     "&#9989;",
                lines:    new[]
                {
                    $"<strong>{approver}</strong> approved your document at Level {level}.",
                    isFinal
                        ? $"<br/>&#127881; All {totalLevels} approval level(s) are now complete. Your document is fully approved."
                        : $"<br/>The document has advanced to Level {level + 1} for further review."
                },
                actionUrl:  null,
                actionLabel:null,
                footerNote: $"Document: {itemName}"
            );
            SendEmail(submitterEmail, subject, body);
        }

        public void NotifyRejected(string submitterEmail, string itemName,
                                   string approver, int level, string comments)
        {
            if (!_enabled) return;
            var subject = $"[Rejected] {itemName} - requires your attention";
            var body = BuildBody(
                title:    "Document Rejected",
                iconColor:"#E84F26",
                icon:     "&#10060;",
                lines:    new[]
                {
                    $"<strong>{approver}</strong> rejected your document at Level {level}.",
                    $"<br/>Reason: <em>{EscapeHtml(comments ?? "No reason provided.")}</em>",
                    "<br/>Please review the feedback and resubmit if required."
                },
                actionUrl:  null,
                actionLabel:null,
                footerNote: $"Document: {itemName}"
            );
            SendEmail(submitterEmail, subject, body);
        }

        public void NotifyDelegated(string newApproverEmail, string itemName,
                                    string delegatedFrom, int level, string reason)
        {
            if (!_enabled) return;
            var subject = $"[Delegated to you] Approval required for {itemName}";
            var body = BuildBody(
                title:    "Approval Delegated to You",
                iconColor:"#00A7E4",
                icon:     "&#128101;",
                lines:    new[]
                {
                    $"<strong>{delegatedFrom}</strong> has delegated their Level {level} approval to you.",
                    $"<br/>Document: <strong>{itemName}</strong>",
                    string.IsNullOrEmpty(reason) ? "" : $"<br/>Note: <em>{EscapeHtml(reason)}</em>",
                    "<br/>Please log in to the DMS to take action."
                },
                actionUrl:  null,
                actionLabel:null,
                footerNote: $"Delegated by {delegatedFrom}"
            );
            SendEmail(newApproverEmail, subject, body);
        }

        public void NotifyEscalated(string escalateToEmail, string itemName,
                                    string approver, int level, DateTime dueDate)
        {
            if (!_enabled) return;
            var subject = $"[Escalation] Overdue approval – {itemName}";
            var body = BuildBody(
                title:    "Approval Overdue - Escalation",
                iconColor:"#FEB900",
                icon:     "&#9888;",
                lines:    new[]
                {
                    $"The approval for <strong>{itemName}</strong> at Level {level} is overdue.",
                    $"<br/>Assigned approver: <strong>{approver}</strong>",
                    $"<br/>Due date was: <strong>{dueDate:dd MMM yyyy}</strong>",
                    "<br/>Please take appropriate action to unblock this workflow."
                },
                actionUrl:  null,
                actionLabel:null,
                footerNote: "This is an automated escalation alert."
            );
            SendEmail(escalateToEmail, subject, body);
        }

        public void NotifyRecalled(string approverEmail, string itemName,
                                   string submitter)
        {
            if (!_enabled) return;
            var subject = $"[Recalled] {itemName} - no action required";
            var body = BuildBody(
                title:    "Document Recalled",
                iconColor:"#6B7280",
                icon:     "&#8635;",
                lines:    new[]
                {
                    $"<strong>{submitter}</strong> has recalled the submission of <strong>{itemName}</strong>.",
                    "<br/>No further action is required from you for this item."
                },
                actionUrl:  null,
                actionLabel:null,
                footerNote: "The workflow has been cancelled."
            );
            SendEmail(approverEmail, subject, body);
        }

        public void NotifyCompleted(string submitterEmail, string itemName, int totalLevels)
        {
            if (!_enabled) return;
            var subject = $"[Completed] {itemName} - fully approved";
            var body = BuildBody(
                title:    "Document Fully Approved",
                iconColor:"#83BB00",
                icon:     "&#127881;",
                lines:    new[]
                {
                    $"Your document <strong>{itemName}</strong> has been approved by all {totalLevels} level(s).",
                    "<br/>The workflow is now complete."
                },
                actionUrl:  null,
                actionLabel:null,
                footerNote: "No further action is required."
            );
            SendEmail(submitterEmail, subject, body);
        }

        // Private helpers

        private string BuildBody(string title, string iconColor, string icon,
                                 string[] lines, string actionUrl, string actionLabel,
                                 string footerNote)
        {
            var sb = new StringBuilder();
            sb.Append($@"
<!DOCTYPE html>
<html>
<head><meta charset='utf-8'/></head>
<body style='margin:0;padding:0;background:#F2F4F8;font-family:Arial,sans-serif;'>
  <table width='100%' cellpadding='0' cellspacing='0'>
    <tr><td align='center' style='padding:32px 16px;'>
      <table width='560' cellpadding='0' cellspacing='0'
             style='background:#fff;border-radius:10px;overflow:hidden;
                    box-shadow:0 2px 12px rgba(0,0,0,.08);'>

        <!-- Header -->
        <tr>
          <td style='background:{iconColor};padding:24px 32px;'>
            <span style='font-size:2rem;'>{icon}</span>
            <span style='color:#fff;font-size:1.15rem;font-weight:700;
                         margin-left:12px;vertical-align:middle;'>{title}</span>
          </td>
        </tr>

        <!-- Body -->
        <tr>
          <td style='padding:28px 32px;color:#374151;font-size:.93rem;line-height:1.7;'>
            {string.Join("", lines)}");

            if (!string.IsNullOrEmpty(actionUrl))
            {
                sb.Append($@"
            <br/><br/>
            <a href='{actionUrl}'
               style='display:inline-block;padding:10px 24px;background:{iconColor};
                      color:#fff;border-radius:6px;text-decoration:none;
                      font-weight:700;font-size:.9rem;'>{actionLabel}</a>");
            }

            sb.Append($@"
          </td>
        </tr>

        <!-- Footer -->
        <tr>
          <td style='background:#F7F8FB;padding:14px 32px;border-top:1px solid #E8EAF0;
                     color:#9CA3AF;font-size:.78rem;'>
            {EscapeHtml(footerNote ?? "")}
            &nbsp;|&nbsp; AppKings DMS &nbsp;|&nbsp; {DateTime.Now:dd MMM yyyy HH:mm}
          </td>
        </tr>

      </table>
    </td></tr>
  </table>
</body>
</html>");
            return sb.ToString();
        }

        private void SendEmail(string to, string subject, string htmlBody)
        {
            if (!_enabled || string.IsNullOrWhiteSpace(to)) return;
            try
            {
                using (var msg = new MailMessage())
                {
                    msg.From       = new MailAddress(_from, _fromName);
                    msg.To.Add(to);
                    msg.Subject    = subject;
                    msg.Body       = htmlBody;
                    msg.IsBodyHtml = true;

                    using (var client = new SmtpClient(_host, _port))
                    {
                        client.EnableSsl = _ssl;
                        if (!string.IsNullOrEmpty(_smtpUser))
                            client.Credentials = string.IsNullOrEmpty(_smtpDomain)
                                ? (ICredentialsByHost)new NetworkCredential(_smtpUser, _smtpPass)
                                : new NetworkCredential(_smtpUser, _smtpPass, _smtpDomain);
                        client.Send(msg);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log and continue - never crash the workflow because of a notification failure.
                System.Diagnostics.Trace.TraceWarning("[NotificationService] Failed to send e-mail to {0}: {1}", to, ex.Message);
            }
        }

        private static string EscapeHtml(string s) =>
            s?.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;") ?? "";
    }
}
