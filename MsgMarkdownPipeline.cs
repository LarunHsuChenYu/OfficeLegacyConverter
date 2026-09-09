using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using MsgReader.Outlook;
using ReverseMarkdown;

namespace OfficeLegacyConverter;

/// <summary>
/// Outlook .msg → Markdown，對齊 mail-analysis 的 raw→extracted 結構：
/// metadata、正文（含 nested quote）、附件清單與抽出檔。
/// </summary>
internal static class MsgMarkdownPipeline
{
    public static readonly string[] SupportedExtensions = [".msg"];

    private static readonly Converter MarkdownConverter = new(new Config
    {
        UnknownTags = Config.UnknownTagsOption.Bypass,
        GithubFlavored = true,
        RemoveComments = true,
        SmartHrefHandling = true,
    });

    public static bool IsSupported(string path) =>
        Path.GetExtension(path).Equals(".msg", StringComparison.OrdinalIgnoreCase);

    public static IEnumerable<string> ExpandPaths(string path)
    {
        if (!Directory.Exists(path))
            return IsSupported(path) ? [path] : [];

        return Directory.EnumerateFiles(path, "*.msg", SearchOption.AllDirectories);
    }

    public static BatchResult ConvertFiles(
        IReadOnlyList<string> paths,
        string outputFolder,
        bool deleteOriginal,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        var outputDir = Path.GetFullPath(outputFolder);
        if (!Directory.Exists(outputDir))
            throw new DirectoryNotFoundException($"輸出資料夾不存在：{outputDir}");

        var total = paths.Count;
        var result = new BatchResult { Total = total };

        for (var i = 0; i < total; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var inputPath = paths[i];
            var current = i + 1;
            progress.Report((current, total, $"開啟 {inputPath}..."));

            try
            {
                var mdPath = ConvertFile(inputPath, outputDir, cancellationToken);
                progress.Report((current, total, $"完成：{Path.GetFileName(inputPath)} → {Path.GetFileName(mdPath)}"));

                if (deleteOriginal)
                {
                    File.Delete(inputPath);
                    progress.Report((current, total, $"已刪除原始檔：{inputPath}"));
                }
            }
            catch (Exception ex)
            {
                result.Failures.Add($"{Path.GetFileName(inputPath)}：{ex.Message}");
                progress.Report((current, total, $"錯誤 - {inputPath}：{ex.Message}"));
            }
        }

        return result;
    }

    public static string ConvertFile(string inputPath, string outputFolder, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var outputDir = Path.GetFullPath(outputFolder);
        var baseName = SanitizeFileName(Path.GetFileNameWithoutExtension(inputPath));
        var mailId = BuildMailId(baseName, inputPath);
        var attachmentsDir = Path.Combine(outputDir, "attachments", mailId);
        var mdPath = Path.Combine(outputDir, $"{mailId}.md");

        PreservePrimaryMailWhenUsingAnalysisLayout(inputPath, outputDir);

        using var message = new Storage.Message(inputPath, FileAccess.Read);
        var attachmentEntries = ExtractAttachments(message, attachmentsDir, cancellationToken);
        var bodyMarkdown = BuildBodyMarkdown(message);
        var markdown = BuildMarkdownDocument(inputPath, message, bodyMarkdown, attachmentEntries, mailId);

        File.WriteAllText(mdPath, markdown, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        return mdPath;
    }

    private static string BuildMarkdownDocument(
        string inputPath,
        Storage.Message message,
        string bodyMarkdown,
        IReadOnlyList<AttachmentEntry> attachments,
        string mailId)
    {
        var subject = NullToEmpty(message.Subject);
        var from = FormatSender(message);
        var sentOn = FormatDate(message.SentOn);
        var to = NullToEmpty(message.GetEmailRecipients(RecipientType.To, false, false));
        var cc = NullToEmpty(message.GetEmailRecipients(RecipientType.Cc, false, false));
        var bcc = NullToEmpty(message.GetEmailRecipients(RecipientType.Bcc, false, false));
        var conversation = NullToEmpty(message.ConversationTopic);
        var messageId = NullToEmpty(message.Id);
        var importance = message.Importance?.ToString() ?? "";

        var sb = new StringBuilder();
        sb.AppendLine("---");
        sb.AppendLine($"source_file: {YamlEscape(Path.GetFileName(inputPath))}");
        sb.AppendLine($"source_path: {YamlEscape(inputPath)}");
        sb.AppendLine($"mail_id: {YamlEscape(mailId)}");
        sb.AppendLine($"converted_date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine("converter: OfficeLegacyConverter MsgMarkdownPipeline");
        sb.AppendLine("document_type: email");
        sb.AppendLine("analysis_status: extraction_only");
        sb.AppendLine($"subject: {YamlEscape(subject)}");
        sb.AppendLine($"from: {YamlEscape(from)}");
        sb.AppendLine($"to: {YamlEscape(to)}");
        sb.AppendLine($"cc: {YamlEscape(cc)}");
        sb.AppendLine($"bcc: {YamlEscape(bcc)}");
        sb.AppendLine($"sent_on: {YamlEscape(sentOn)}");
        sb.AppendLine($"conversation_topic: {YamlEscape(conversation)}");
        sb.AppendLine($"internet_message_id: {YamlEscape(messageId)}");
        if (!string.IsNullOrWhiteSpace(importance))
            sb.AppendLine($"importance: {YamlEscape(importance)}");
        sb.AppendLine($"attachment_count: {attachments.Count}");
        sb.AppendLine("---");
        sb.AppendLine();
        sb.AppendLine($"# {EscapeHeading(subject)}");
        sb.AppendLine();
        sb.AppendLine("| 欄位 | 內容 |");
        sb.AppendLine("|---|---|");
        sb.AppendLine($"| 日期 | {EscapeTable(sentOn)} |");
        sb.AppendLine($"| 寄件 | {EscapeTable(from)} |");
        sb.AppendLine($"| 收件 | {EscapeTable(to)} |");
        if (!string.IsNullOrWhiteSpace(cc))
            sb.AppendLine($"| 副本 | {EscapeTable(cc)} |");
        if (!string.IsNullOrWhiteSpace(bcc))
            sb.AppendLine($"| 密件副本 | {EscapeTable(bcc)} |");
        sb.AppendLine($"| 主旨 | {EscapeTable(subject)} |");
        sb.AppendLine($"| Thread | {EscapeTable(string.IsNullOrWhiteSpace(conversation) ? subject : conversation)} |");
        if (!string.IsNullOrWhiteSpace(messageId))
            sb.AppendLine($"| Message-ID | {EscapeTable(messageId)} |");
        sb.AppendLine();
        sb.AppendLine("## 正文");
        sb.AppendLine();
        sb.AppendLine(string.IsNullOrWhiteSpace(bodyMarkdown) ? "_（無正文）_" : bodyMarkdown.TrimEnd());
        sb.AppendLine();
        sb.AppendLine("## 附件");
        sb.AppendLine();

        if (attachments.Count == 0)
        {
            sb.AppendLine("_（無附件）_");
        }
        else
        {
            foreach (var attachment in attachments)
            {
                var linkTarget = attachment.RelativePath.Replace('\\', '/');
                var kind = attachment.IsEmbeddedMessage ? "內嵌郵件" : "檔案";
                sb.AppendLine($"- [{attachment.FileName}]({linkTarget}) （{kind}{(string.IsNullOrWhiteSpace(attachment.ContentId) ? "" : $", CID: {attachment.ContentId}")}）");
            }
        }

        sb.AppendLine();
        sb.AppendLine("## Mail analysis 交接");
        sb.AppendLine();
        sb.AppendLine("- 狀態：本檔只完成郵件與附件擷取，尚未完成 claim ledger、thread 完整性與附件內容核對。");
        sb.AppendLine("- 證據：定案與 pending 必須回到原始 `.msg`、本檔正文及客戶附件，不得只引用後續分析文件。");
        sb.AppendLine("- 下一步：依 `/mail-analysis` 先逐句處理最上層正文，再讀 nested quote，最後開啟並核對附件。");
        sb.AppendLine();
        return sb.ToString();
    }

    private static void PreservePrimaryMailWhenUsingAnalysisLayout(string inputPath, string outputDir)
    {
        if (!Path.GetFileName(outputDir.TrimEnd(
                Path.DirectorySeparatorChar,
                Path.AltDirectorySeparatorChar))
            .Equals("extracted", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var mailRoot = Directory.GetParent(outputDir)?.FullName;
        if (string.IsNullOrEmpty(mailRoot))
            return;

        var rawDir = Path.Combine(mailRoot, "raw");
        Directory.CreateDirectory(rawDir);
        var destination = Path.Combine(rawDir, Path.GetFileName(inputPath));

        if (Path.GetFullPath(inputPath).Equals(
                Path.GetFullPath(destination),
                StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (File.Exists(destination))
        {
            var sourceInfo = new FileInfo(inputPath);
            var destinationInfo = new FileInfo(destination);
            if (sourceInfo.Length == destinationInfo.Length)
                return;

            var stamp = File.GetLastWriteTime(inputPath)
                .ToString("yyyyMMdd_HHmm", CultureInfo.InvariantCulture);
            destination = Path.Combine(rawDir, $"{stamp}_{Path.GetFileName(inputPath)}");
        }

        File.Copy(inputPath, destination, overwrite: false);
    }

    private static string BuildBodyMarkdown(Storage.Message message)
    {
        try
        {
            var html = message.BodyHtml;
            if (!string.IsNullOrWhiteSpace(html))
            {
                try
                {
                    var md = MarkdownConverter.Convert(html);
                    return NormalizeQuotes(md);
                }
                catch
                {
                    // fall through to plain text
                }
            }
        }
        catch
        {
            // BodyHtml 產生失敗時改走純文字
        }

        try
        {
            var text = message.BodyText;
            if (!string.IsNullOrWhiteSpace(text))
                return NormalizeQuotes(text);
        }
        catch
        {
            // ignore
        }

        try
        {
            var rtf = message.BodyRtf;
            if (!string.IsNullOrWhiteSpace(rtf))
                return NormalizeQuotes(StripRtfRoughly(rtf));
        }
        catch
        {
            // ignore
        }

        return "";
    }

    private static string NormalizeQuotes(string body)
    {
        var normalized = body.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var sb = new StringBuilder();

        foreach (var rawLine in lines)
        {
            var line = rawLine;

            // Outlook 常見引用前綴轉成 Markdown blockquote，保留 nested depth
            var quoteDepth = 0;
            while (true)
            {
                var trimmedStart = line.TrimStart();
                if (trimmedStart.StartsWith("> ", StringComparison.Ordinal))
                {
                    quoteDepth++;
                    line = trimmedStart[2..];
                    continue;
                }

                if (trimmedStart.StartsWith("|", StringComparison.Ordinal)
                    && !trimmedStart.StartsWith("|---", StringComparison.Ordinal))
                {
                    // 部分純文字回覆用 | 當引用
                    quoteDepth++;
                    line = trimmedStart[1..].TrimStart();
                    continue;
                }

                break;
            }

            if (quoteDepth > 0)
                sb.Append(string.Concat(Enumerable.Repeat("> ", quoteDepth)));

            sb.AppendLine(line);
        }

        return sb.ToString().TrimEnd() + "\n";
    }

    private static List<AttachmentEntry> ExtractAttachments(
        Storage.Message message,
        string attachmentsDir,
        CancellationToken cancellationToken)
    {
        var entries = new List<AttachmentEntry>();
        if (message.Attachments is null || message.Attachments.Count == 0)
            return entries;

        Directory.CreateDirectory(attachmentsDir);
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var attachmentObj in message.Attachments)
        {
            cancellationToken.ThrowIfCancellationRequested();

            switch (attachmentObj)
            {
                case Storage.Attachment fileAttachment:
                {
                    var fileName = ResolveAttachmentFileName(fileAttachment, usedNames);
                    var savePath = Path.Combine(attachmentsDir, fileName);
                    SaveAttachmentData(fileAttachment, savePath);

                    entries.Add(new AttachmentEntry(
                        fileName,
                        ToAttachmentsRelativePath(attachmentsDir, savePath),
                        fileAttachment.ContentId,
                        IsEmbeddedMessage: false));
                    break;
                }
                case Storage.Message embedded:
                {
                    var subject = string.IsNullOrWhiteSpace(embedded.Subject) ? "embedded" : embedded.Subject!;
                    var fileName = EnsureUniqueName(SanitizeFileName(subject) + ".msg", usedNames);
                    var savePath = Path.Combine(attachmentsDir, fileName);
                    try
                    {
                        embedded.Save(savePath);
                    }
                    catch
                    {
                        // 若無法另存內嵌 MSG，仍列於清單
                        File.WriteAllText(
                            Path.ChangeExtension(savePath, ".txt"),
                            $"內嵌郵件：{subject}\n寄件：{FormatSender(embedded)}\n日期：{FormatDate(embedded.SentOn)}",
                            Encoding.UTF8);
                        fileName = Path.GetFileName(Path.ChangeExtension(savePath, ".txt"));
                        savePath = Path.Combine(attachmentsDir, fileName);
                    }

                    entries.Add(new AttachmentEntry(
                        fileName,
                        ToAttachmentsRelativePath(attachmentsDir, savePath),
                        ContentId: null,
                        IsEmbeddedMessage: true));
                    break;
                }
            }
        }

        return entries;
    }

    private static void SaveAttachmentData(Storage.Attachment attachment, string savePath)
    {
        if (attachment.Data is { Length: > 0 } data)
        {
            File.WriteAllBytes(savePath, data);
            return;
        }

        // 部分版本以串流／暫存提供；退回寫入空檔並標註
        File.WriteAllBytes(savePath, []);
    }

    private static string ResolveAttachmentFileName(Storage.Attachment attachment, HashSet<string> usedNames)
    {
        var preferred = attachment.FileName
            ?? (!string.IsNullOrWhiteSpace(attachment.ContentId) ? attachment.ContentId + ".bin" : null)
            ?? "attachment.bin";

        return EnsureUniqueName(SanitizeFileName(preferred), usedNames);
    }

    private static string ToAttachmentsRelativePath(string attachmentsDir, string savePath)
    {
        // attachments/{mailId}/file → relative from output folder (= parent of attachments)
        var outputRoot = Directory.GetParent(attachmentsDir)?.Parent?.FullName
            ?? Path.GetDirectoryName(attachmentsDir)
            ?? ".";
        return Path.GetRelativePath(outputRoot, savePath);
    }

    private static string BuildMailId(string baseName, string inputPath)
    {
        var stamp = File.GetLastWriteTime(inputPath).ToString("yyyyMMdd_HHmm", CultureInfo.InvariantCulture);
        var shortName = baseName.Length > 48 ? baseName[..48] : baseName;
        return $"{stamp}_{shortName}";
    }

    private static string FormatSender(Storage.Message message)
    {
        try
        {
            var formatted = NullToEmpty(message.GetEmailSender(html: false, convertToHref: false));
            if (!string.IsNullOrWhiteSpace(formatted))
                return formatted;
        }
        catch
        {
            // fall through
        }

        try
        {
            var sender = message.Sender;
            if (sender is null)
                return "";

            var display = NullToEmpty(sender.DisplayName);
            var email = NullToEmpty(sender.Email);
            if (!string.IsNullOrWhiteSpace(display) && !string.IsNullOrWhiteSpace(email))
                return $"{display} <{email}>";
            return string.IsNullOrWhiteSpace(display) ? email : display;
        }
        catch
        {
            return "";
        }
    }

    private static string FormatDate(DateTimeOffset? value) =>
        value is { } dto && dto != DateTimeOffset.MinValue
            ? dto.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
            : "";

    private static string EnsureUniqueName(string fileName, HashSet<string> usedNames)
    {
        if (usedNames.Add(fileName))
            return fileName;

        var name = Path.GetFileNameWithoutExtension(fileName);
        var ext = Path.GetExtension(fileName);
        for (var i = 2; ; i++)
        {
            var candidate = $"{name}_{i}{ext}";
            if (usedNames.Add(candidate))
                return candidate;
        }
    }

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? "untitled" : cleaned;
    }

    private static string YamlEscape(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "\"\"";

        if (value.IndexOfAny([':', '#', '{', '}', '[', ']', ',', '&', '*', '!', '|', '>', '\'', '"', '%', '@', '`']) >= 0
            || value.Contains('\n')
            || value.StartsWith(' ')
            || value.EndsWith(' '))
        {
            return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n") + "\"";
        }

        return value;
    }

    private static string EscapeTable(string value) =>
        value.Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");

    private static string EscapeHeading(string value) =>
        string.IsNullOrWhiteSpace(value) ? "(無主旨)" : value.Replace("\r", " ").Replace("\n", " ");

    private static string NullToEmpty(string? value) => value?.Trim() ?? "";

    private static string StripRtfRoughly(string rtf)
    {
        var withoutGroups = Regex.Replace(rtf, @"\\'[0-9a-fA-F]{2}", " ");
        withoutGroups = Regex.Replace(withoutGroups, @"\\[a-zA-Z]+-?\d* ?", " ");
        withoutGroups = withoutGroups.Replace("{", " ").Replace("}", " ");
        return Regex.Replace(withoutGroups, @"\s+", " ").Trim();
    }

    private sealed record AttachmentEntry(
        string FileName,
        string RelativePath,
        string? ContentId,
        bool IsEmbeddedMessage);
}
