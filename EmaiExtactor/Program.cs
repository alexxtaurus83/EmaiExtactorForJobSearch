using MailKit.Net.Imap;
using MailKit;
using MailKit.Search;
using MimeKit;
using System;
using System.IO;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using System.Collections.Generic;
using System.Linq;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.Net;
using System.Text.Json;

public class SimplifiedEmailEntry {
    public string DateTime { get; set; }
    public string From { get; set; }
    public string Subject { get; set; }
    public string Body { get; set; }
    public string EmailLink { get; set; }
    public string CompanyName { get; set; }
    public bool IsDirectCompany { get; set; }
    public string JobTitle { get; set; }
    public string ApplicationStatus { get; set; }

    [CsvHelper.Configuration.Attributes.Ignore]
    public DateTimeOffset ParsedDate { get; set; }
    [CsvHelper.Configuration.Attributes.Ignore]
    public string FullTextForAnalysis { get; set; }
}

class Program {
    private static AppConfig _config;

    static void Main(string[] args) {
        // Load configuration from appsettings.json
        try {
            var configJson = File.ReadAllText("appsettings.json");
            _config = JsonSerializer.Deserialize<AppConfig>(configJson);
        } catch (Exception ex) {
            Console.WriteLine($"Error loading configuration: {ex.Message}");
            return;
        }

        Console.WriteLine("Fetching emails...");
        var rawEmailData = GetRawEmails();

        Console.WriteLine("Processing emails with simplified logic...");
        var simplifiedEmails = ProcessEmailsSimplified(rawEmailData);

        SaveSimplifiedToCsv(simplifiedEmails, _config.OutputSettings.CsvFilename);
        Console.WriteLine($"Simplified emails extracted successfully to {_config.OutputSettings.CsvFilename}!");
    }

    static string ConvertHtmlToPlainText(string html) {
        if (string.IsNullOrEmpty(html)) return string.Empty;
        var doc = new HtmlDocument();
        doc.LoadHtml(html);
        var plainText = WebUtility.HtmlDecode(doc.DocumentNode.InnerText);
        return Regex.Replace(plainText, @"\s+", " ").Trim();
    }

    static List<string[]> GetRawEmails() {
        var emailList = new List<string[]>();
        using (var client = new ImapClient()) {
            try {
                client.Connect(_config.ImapSettings.ImapServer, _config.ImapSettings.ImapPort, true);
                client.Authenticate(_config.ImapSettings.EmailUser, _config.ImapSettings.Password);

                var targetLabel = _config.ImapSettings.EmailLabel;
                var targetFolder = client.GetFolders(client.PersonalNamespaces[0]).FirstOrDefault(f => f.Name.Equals(targetLabel, StringComparison.OrdinalIgnoreCase));

                if (targetFolder == null) {
                    Console.WriteLine($"Label '{targetLabel}' not found!"); return emailList;
                }

                targetFolder.Open(FolderAccess.ReadOnly);
                var allUids = targetFolder.Search(SearchQuery.All);
                Console.WriteLine($"Found {allUids.Count} emails in label '{targetLabel}'. Fetching...");

                int idx = 0;
                foreach (var uid in allUids) {
                    try {
                        var message = targetFolder.GetMessage(uid);
                        emailList.Add(new string[] {
                            message.Date.ToString("o"),
                            message.From.ToString(),
                            message.Subject ?? "",
                            message.HtmlBody != null ? ConvertHtmlToPlainText(message.HtmlBody) : (message.TextBody ?? ""),
                            $"https://mail.google.com/mail/u/0/#label/{WebUtility.UrlEncode(targetLabel)}/{uid}"
                        });
                        idx++;
                        if (idx % 20 == 0) Console.WriteLine($"Fetched {idx}/{allUids.Count} emails...");
                    } catch (Exception ex) { Console.WriteLine($"Error fetching message UID {uid}: {ex.Message}"); }
                }
            } catch (Exception ex) { Console.WriteLine($"IMAP client error: {ex.Message}"); } finally { if (client.IsConnected) client.Disconnect(true); }
        }
        return emailList;
    }

    static List<SimplifiedEmailEntry> ProcessEmailsSimplified(List<string[]> rawEmails) {
        var processedList = new List<SimplifiedEmailEntry>();
        var knownAtsDomains = new HashSet<string>(_config.DomainLists.KnownAtsDomains, StringComparer.OrdinalIgnoreCase);
        var knownJobBoardDomains = new HashSet<string>(_config.DomainLists.KnownJobBoardDomains, StringComparer.OrdinalIgnoreCase);

        foreach (var rawEmailArray in rawEmails) {
            var entry = new SimplifiedEmailEntry {
                DateTime = rawEmailArray[0],
                From = rawEmailArray[1],
                Subject = rawEmailArray[2],
                Body = rawEmailArray[3],
                EmailLink = rawEmailArray[4]
            };

            try {
                entry.ParsedDate = DateTimeOffset.Parse(entry.DateTime, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
            } catch { entry.ParsedDate = DateTimeOffset.MinValue; }

            entry.FullTextForAnalysis = (entry.Subject + " " + entry.Body).ToLowerInvariant();

            entry.CompanyName = ExtractCompanyName(entry.From, entry.Subject, entry.Body, entry.FullTextForAnalysis, knownAtsDomains, knownJobBoardDomains);
            entry.JobTitle = ExtractJobTitle(entry.Subject, entry.Body, entry.CompanyName);
            entry.IsDirectCompany = DetermineDirectCompany(entry.From, entry.CompanyName, entry.FullTextForAnalysis, knownAtsDomains, knownJobBoardDomains);
            entry.ApplicationStatus = DetermineApplicationStatus(entry.FullTextForAnalysis, entry.Subject.ToLowerInvariant());

            processedList.Add(entry);
        }
        return processedList.OrderBy(e => e.ParsedDate).ToList();
    }

    static string ExtractCompanyName(string fromText, string subject, string body, string fullTextLower, HashSet<string> knownAtsDomains, HashSet<string> knownJobBoardDomains) {
        Match match;

        foreach (var pattern in _config.ExtractionPatterns.CompanyNamePatterns) {
            match = Regex.Match(subject, pattern, RegexOptions.IgnoreCase);
            if (match.Success) return CleanCompanyName(match.Groups[1].Value);
        }

        // Check body patterns as well
        foreach (var pattern in _config.ExtractionPatterns.CompanyNamePatterns) {
            match = Regex.Match(body, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
            if (match.Success && match.Groups[1].Length > 0 && IsLikelyCompanyName(match.Groups[1].Value)) {
                return CleanCompanyName(match.Groups[1].Value);
            }
        }

        foreach (var cName in _config.KeywordLists.SpecificCompanyKeywords) {
            if (fullTextLower.Contains(cName.ToLowerInvariant())) return cName;
        }

        var parsedFrom = MimeKit.MailboxAddress.Parse(fromText);
        if (parsedFrom != null) {
            if (!string.IsNullOrEmpty(parsedFrom.Name) && !parsedFrom.Name.Contains("@") && IsLikelyCompanyName(parsedFrom.Name)) {
                return CleanCompanyName(parsedFrom.Name);
            }
            var domainParts = parsedFrom.Address.Split('@').LastOrDefault()?.Split('.');
            if (domainParts != null && domainParts.Length >= 2) {
                var potentialCompany = domainParts[domainParts.Length - 2];
                if (IsLikelyCompanyName(potentialCompany) && !knownAtsDomains.Any(ats => parsedFrom.Address.ToLowerInvariant().Contains(ats)) && !knownJobBoardDomains.Any(jb => parsedFrom.Address.ToLowerInvariant().Contains(jb)))
                    return CleanCompanyName(potentialCompany);
            }
        }

        return "Unknown";
    }

    static string ExtractJobTitle(string subject, string body, string companyName) {
        string jobTitle = "Unknown";
        string[] textsToSearch = { subject, body };

        var jobTitlePatterns = _config.ExtractionPatterns.JobTitlePatterns.Select(p => new Regex(p, RegexOptions.IgnoreCase)).ToList();

        foreach (string textSource in textsToSearch) {
            if (string.IsNullOrEmpty(textSource)) continue;

            foreach (var pattern in jobTitlePatterns) {
                Regex currentPattern = pattern;
                // Inject company name if pattern supports it, to help avoid capturing company name as job title
                if (pattern.ToString().Contains("#{companyName}") && companyName != "Unknown" && !string.IsNullOrWhiteSpace(companyName)) {
                    currentPattern = new Regex(pattern.ToString().Replace("#{companyName}", Regex.Escape(companyName)), RegexOptions.IgnoreCase);
                }

                Match titleMatch = currentPattern.Match(textSource);
                if (titleMatch.Success) {
                    var potentialTitle = titleMatch.Groups[1].Value.Trim();

                    // Cleanup Logic
                    potentialTitle = Regex.Replace(potentialTitle, @"^[A-Z0-9-]+\s+(?=[A-Za-z])", "").Trim();
                    potentialTitle = Regex.Replace(potentialTitle, @"\s*\(?(?:Remote|USA|US|REMOTE|Hybrid|Open|Target|Req ID.*)\)?$", "", RegexOptions.IgnoreCase).Trim();
                    potentialTitle = Regex.Replace(potentialTitle, @"(?:\s+role|\s+position|\s+opening|\s+job|\s+opportunity)$", "", RegexOptions.IgnoreCase).Trim();
                    potentialTitle = Regex.Replace(potentialTitle, @"\s*\[\w+\]$", "").Trim();

                    if (!string.IsNullOrWhiteSpace(companyName) && companyName != "Unknown") {
                        potentialTitle = Regex.Replace(potentialTitle, $@"\s+at\s+{Regex.Escape(companyName)}\b.*$", "", RegexOptions.IgnoreCase).Trim();
                    }
                    potentialTitle = Regex.Replace(potentialTitle, @"\s+at\s+[\w\s.&,'-]+$", "", RegexOptions.IgnoreCase).Trim();
                    potentialTitle = Regex.Replace(potentialTitle, @"\bSr\.?\b", "Senior", RegexOptions.IgnoreCase);
                    potentialTitle = potentialTitle.Trim(' ', ',', '-', '–', '(', ')', '‘', '’', '"', '\'', ':');

                    if (Regex.IsMatch(potentialTitle, @"^(thank you|application|resume|we have received|your interest|for the|opening for|details for|the role of the)\b", RegexOptions.IgnoreCase)) {
                        potentialTitle = "Unknown";
                    }

                    if (potentialTitle.Length > 3 && potentialTitle.Length < 120 &&
                        !potentialTitle.Equals(companyName, StringComparison.OrdinalIgnoreCase) &&
                        !IsLikelyCompanyName(potentialTitle) &&
                        Regex.IsMatch(potentialTitle, @"\w.*\w") &&
                        potentialTitle.Split(' ').Length <= 10) {
                        jobTitle = potentialTitle;
                        return jobTitle; // Exit once a plausible title is found
                    }
                }
            }
        }
        return jobTitle;
    }

    static bool DetermineDirectCompany(string fromText, string companyName, string fullTextLower, HashSet<string> knownAtsDomains, HashSet<string> knownJobBoardDomains) {
        if (companyName == "Unknown") return false;

        var parsedFrom = MimeKit.MailboxAddress.Parse(fromText);
        if (parsedFrom == null || string.IsNullOrEmpty(parsedFrom.Address)) return false;

        string senderDomain = parsedFrom.Address.Split('@').LastOrDefault()?.ToLowerInvariant();
        if (string.IsNullOrEmpty(senderDomain)) return false;

        string cleanedCompanyNameForDomain = Regex.Replace(companyName.ToLowerInvariant(), @"\s+(inc|llc|ltd|corp|group|labs|solutions|systems|corporation)$", "").Replace(".", "").Replace(",", "").Split(' ')[0];
        if (senderDomain.Contains(cleanedCompanyNameForDomain) && !IsGenericDomain(senderDomain)) {
            return true;
        }

        if (knownAtsDomains.Any(ats => senderDomain.Contains(ats))) {
            return companyName != "Unknown";
        }

        if (fullTextLower.Contains(companyName.ToLowerInvariant() + " talent acquisition") ||
            fullTextLower.Contains(companyName.ToLowerInvariant() + " recruiting team") ||
            fullTextLower.Contains(companyName.ToLowerInvariant() + " hiring team") ||
            fullTextLower.Contains("the " + companyName.ToLowerInvariant() + " team")) {
            if (IsGenericDomain(senderDomain)) return true;
        }

        if (knownJobBoardDomains.Any(board => senderDomain.Contains(board))) {
            if (senderDomain.Contains("linkedin.com")) {
                return fullTextLower.Contains("application for") && fullTextLower.Contains(companyName.ToLowerInvariant());
            }
            return false;
        }

        return companyName != "Unknown";
    }

    static string DetermineApplicationStatus(string fullTextLower, string subjectLower) {
        if (_config.KeywordLists.RejectionKeywords.Any(kw => fullTextLower.Contains(kw))) {
            return "Rejection Reply";
        }

        if (_config.KeywordLists.OriginalApplicationKeywords.Any(kw => subjectLower.Contains(kw) || fullTextLower.Contains(kw))) {
            if (subjectLower.Contains("good move") && fullTextLower.Contains("shopify")) return "Original Application";
            if (fullTextLower.Contains("thank you for your interest in") && (fullTextLower.Contains("applying for") || fullTextLower.Contains("application for"))) return "Original Application";
            return "Original Application";
        }

        return "Other";
    }

    private static string CleanCompanyName(string name) {
        if (string.IsNullOrWhiteSpace(name)) return "Unknown";
        var cleaned = name.Trim();
        cleaned = Regex.Replace(cleaned, @"\s+(Inc\.?|LLC|L\.P\.?|Ltd\.?|Corp\.?|Corporation|Group|Labs|Solutions|Systems|Company|Careers)$", "", RegexOptions.IgnoreCase);
        cleaned = cleaned.Replace("\"", "").Replace("careers", "", StringComparison.OrdinalIgnoreCase).Trim();
        if (cleaned.StartsWith("The ", StringComparison.OrdinalIgnoreCase)) cleaned = cleaned.Substring(4);
        cleaned = cleaned.TrimEnd('.', ',', '!', '?', ':', ';', ' ');
        return string.IsNullOrWhiteSpace(cleaned) || cleaned.Length < 2 ? "Unknown" : cleaned;
    }

    private static bool IsLikelyCompanyName(string name) {
        if (string.IsNullOrWhiteSpace(name)) return false;
        var lower = name.ToLower();
        if (lower.Contains("dear") || lower.Contains("hello") || lower.Contains("hi ") ||
            lower.Length < 2 || name.Contains("@") || name.Contains("<") || name.Contains(">") ||
            new[] { "team", "sincerely", "regards", "best", "thank", "thanks", "cheers", "talent acquisition", "recruiting team", "hiring team", "auto-reply", "notification", "support", "system", "do-not-reply", "noreply", "no-reply" }.Contains(lower) ||
            Regex.IsMatch(name, @"^\d+$") ||
            name.Split(' ').Length > 5)
            return false;
        return Regex.IsMatch(name, @"^[A-Za-z0-9\s\.'&/–-]+$", RegexOptions.IgnoreCase);
    }

    private static bool IsGenericDomain(string domain) {
        if (string.IsNullOrEmpty(domain)) return true;
        var genericDomains = new HashSet<string>(_config.DomainLists.GenericDomains, StringComparer.OrdinalIgnoreCase);
        return genericDomains.Contains(domain);
    }

    static void SaveSimplifiedToCsv(List<SimplifiedEmailEntry> data, string filename) {
        using (var writer = new StreamWriter(filename))
        using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture) { ShouldQuote = (args) => true })) {
            csv.WriteRecords(data);
        }
    }
}
