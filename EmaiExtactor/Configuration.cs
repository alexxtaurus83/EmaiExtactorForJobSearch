using System.Collections.Generic;
using System.Text.Json.Serialization;

// Main container class for all configuration settings
public class AppConfig {
    [JsonPropertyName("imapSettings")]
    public ImapSettings ImapSettings { get; set; }

    [JsonPropertyName("extractionPatterns")]
    public ExtractionPatterns ExtractionPatterns { get; set; }

    [JsonPropertyName("domainLists")]
    public DomainLists DomainLists { get; set; }

    [JsonPropertyName("keywordLists")]
    public KeywordLists KeywordLists { get; set; }

    [JsonPropertyName("outputSettings")]
    public OutputSettings OutputSettings { get; set; }
}

// Holds IMAP server and user credential settings
public class ImapSettings {
    [JsonPropertyName("emailUser")]
    public string EmailUser { get; set; }

    [JsonPropertyName("password")]
    public string Password { get; set; } // It's better to use secure storage, but for config, it's here

    [JsonPropertyName("imapServer")]
    public string ImapServer { get; set; }

    [JsonPropertyName("imapPort")]
    public int ImapPort { get; set; }

    [JsonPropertyName("emailLabel")]
    public string EmailLabel { get; set; }
}

// Contains all the regular expressions for data extraction
public class ExtractionPatterns {
    [JsonPropertyName("companyNamePatterns")]
    public List<string> CompanyNamePatterns { get; set; }

    [JsonPropertyName("jobTitlePatterns")]
    public List<string> JobTitlePatterns { get; set; }
}

// Contains lists of known domains for classification
public class DomainLists {
    [JsonPropertyName("knownAtsDomains")]
    public List<string> KnownAtsDomains { get; set; }

    [JsonPropertyName("knownJobBoardDomains")]
    public List<string> KnownJobBoardDomains { get; set; }

    [JsonPropertyName("genericDomains")]
    public List<string> GenericDomains { get; set; }
}

// Contains lists of keywords for parsing and classification
public class KeywordLists {
    [JsonPropertyName("specificCompanyKeywords")]
    public List<string> SpecificCompanyKeywords { get; set; }

    [JsonPropertyName("rejectionKeywords")]
    public List<string> RejectionKeywords { get; set; }

    [JsonPropertyName("originalApplicationKeywords")]
    public List<string> OriginalApplicationKeywords { get; set; }
}

// Defines settings for the output file
public class OutputSettings {
    [JsonPropertyName("csvFilename")]
    public string CsvFilename { get; set; }
}
