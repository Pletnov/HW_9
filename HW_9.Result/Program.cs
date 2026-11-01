using System.Text.RegularExpressions;
using System.Net;
using MailKit;

Console.WriteLine("IMAP");

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.InputEncoding = System.Text.Encoding.UTF8;

var mailServer = "127.0.0.1";
var mailPort = 143;
var userName = "user2@test.local";
var password = "11111";

using var client = new MailKit.Net.Imap.ImapClient();
using var httpClient = new HttpClient();

List<string> messages = new List<string>();
List<string> downloaded = new List<string>();

var urlRegex = new Regex(@"https?://[^\s<>""']+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
var downloads = Directory.CreateDirectory("downloads");

while (true)
{
    try
    {
        client.Connect(mailServer, mailPort, false);
        client.Authenticate(userName, password);
        var inbox = client.Inbox;
        inbox.Open(FolderAccess.ReadWrite);

        for (int i = 0; i < inbox.Count; i++)
        {
            var message = inbox.GetMessage(i);
            string messageId = message.MessageId ?? $"{i}_{message.Date}";
            if (messages.Contains(messageId))
            {
                continue;
            }
            messages.Add(messageId);
            Console.WriteLine($"Message {i + 1}:");
            Console.WriteLine($"From: {message.From}");
            Console.WriteLine($"Theme: {message.Subject}");
            Console.WriteLine($"Date: {message.Date}");
            string body = message.TextBody ?? message.HtmlBody ?? string.Empty;
            var matches = urlRegex.Matches(body);
            foreach (Match match in matches)
            {
                string url = match.Value.TrimEnd(
                    '.', ',', ')', ']', ';'
                );
                Console.WriteLine();
                Console.WriteLine($"Found URL: {url}");
                if (url.Contains("google.com/imgres"))
                {
                    string[] parameters = url.Split('?').Last().Split('&');
                    foreach (string parameter in parameters)
                    {
                        if (parameter.StartsWith("imgurl="))
                        {
                            url = parameter.Substring(7);
                            break;
                        }
                    }
                }
                if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                {
                    continue;
                }
                if (downloaded.Contains(url))
                {
                    Console.WriteLine("Already downloaded");
                    continue;
                }
                try
                {
                    Console.WriteLine("Downloading...");
                    using var response = await httpClient.GetAsync(url);
                    string fileName = url.Split('?')[0].Split('/').Last();
                    string contentType = response.Content.Headers.ContentType?.MediaType ?? "";
                    if (string.IsNullOrWhiteSpace(fileName))
                    {
                        fileName = "file_" + DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
                    }
                    string extension = "";
                    if (contentType == "image/jpeg")
                        extension = ".jpg";
                    else if (contentType == "image/png")
                        extension = ".png";
                    else if (contentType == "image/webp")
                        extension = ".webp";
                    else if (contentType == "image/gif")
                        extension = ".gif";
                    else if (contentType == "application/pdf")
                        extension = ".pdf";
                    else if (contentType == "application/zip")
                        extension = ".zip";
                    if (string.IsNullOrEmpty(Path.GetExtension(fileName)))
                    {
                        fileName += extension;
                    }

                    string fullPath = Path.Combine(downloads.FullName, fileName);
                    if (File.Exists(fullPath))
                    {
                        Console.WriteLine("File already exists");
                        downloaded.Add(url);
                        continue;
                    }
                    long? totalBytes = response.Content.Headers.ContentLength;

                    using var input = await response.Content.ReadAsStreamAsync();
                    using var output = File.Create(fullPath);

                    byte[] buffer = new byte[8192];

                    long downloadedBytes = 0;
                    int bytesRead;

                    while ((bytesRead = await input.ReadAsync(buffer, 0, buffer.Length)) > 0)
                    {
                        await output.WriteAsync(buffer, 0, bytesRead);
                        downloadedBytes += bytesRead;
                        if (totalBytes.Value > 0)
                        {
                            double percent = (double)downloadedBytes / totalBytes.Value * 100;
                            Console.Write($"\rProgress: {percent:F1}%");
                        }
                        else
                        {
                            Console.Write($"\rDownloaded: {downloadedBytes} bytes");
                        }
                    }
                    Console.WriteLine();
                    downloaded.Add(url);
                    Console.WriteLine($"File successfully downloaded: {fullPath}");
                    inbox.AddFlags(i, MailKit.MessageFlags.Deleted, true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message}");
                }
            }
            Console.WriteLine();
        }
        inbox.Expunge();
        client.Disconnect(true);

        Console.WriteLine("------------------------------");
        Console.WriteLine("Waiting 30 seconds...");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"IMAP error: {ex.Message}");
        client.Disconnect(true);
    }
    await Task.Delay(30000);
}
