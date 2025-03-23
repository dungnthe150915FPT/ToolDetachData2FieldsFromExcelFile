using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using System.IO.Compression;
using System.Net.Mail;
using System.Net;

public class IndexModel : PageModel
{
    [BindProperty]
    public IFormFile UploadedFile { get; set; }

    [BindProperty]
    public string RecipientEmail { get; set; }

    public List<string> DownloadLinks { get; set; } = new();
    public string StatusMessage { get; set; }

    public async Task<IActionResult> OnPostAsync()
    {
        if (UploadedFile == null || UploadedFile.Length == 0)
            return Page();

        var filePath = Path.Combine("wwwroot", "uploads", Path.GetRandomFileName() + Path.GetExtension(UploadedFile.FileName));
        Directory.CreateDirectory(Path.GetDirectoryName(filePath));

        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await UploadedFile.CopyToAsync(stream);
        }

        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.First();
        var rowCount = worksheet.Dimension.Rows;

        var records = new List<(string Username, string Org)>();
        for (int row = 2; row <= rowCount; row++)
        {
            var username = worksheet.Cells[row, 1].Text.Trim();
            var org = worksheet.Cells[row, 2].Text.Trim();
            if (!string.IsNullOrWhiteSpace(username) && !string.IsNullOrWhiteSpace(org))
                records.Add((username, org));
        }

        var grouped = records.GroupBy(r => r.Org);
        foreach (var group in grouped)
        {
            var newFile = new ExcelPackage();
            var sheet = newFile.Workbook.Worksheets.Add("Users");
            sheet.Cells[1, 1].Value = "USERNAME";
            sheet.Cells[1, 2].Value = "ORG_CODE_NAME_BDT";

            int row = 2;
            foreach (var item in group)
            {
                sheet.Cells[row, 1].Value = item.Username;
                sheet.Cells[row, 2].Value = item.Org;
                row++;
            }

            var safeName = string.Join("_", group.Key.Split(Path.GetInvalidFileNameChars()));
            var outputPath = Path.Combine("wwwroot", "downloads", $"{safeName}.xlsx");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
            await newFile.SaveAsAsync(new FileInfo(outputPath));

            DownloadLinks.Add("~/downloads/" + Path.GetFileName(outputPath));
        }

        var downloadFolder = Path.Combine("wwwroot", "downloads");
        var zipPath = Path.Combine(Path.GetTempPath(), $"AllFiles_{Guid.NewGuid()}.zip");
        var filesToZip = Directory.GetFiles(downloadFolder, "*.xlsx");

        using (var zipStream = new FileStream(zipPath, FileMode.Create))
        {
            using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var file in filesToZip)
                {
                    archive.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(RecipientEmail))
        {
            try
            {
                var mail = "dungdev224@gmail.com";
                var appPassword = "tnjgrgrmsvbvtvko";

                var message = new MailMessage();
                message.To.Add(RecipientEmail);
                message.From = new MailAddress(mail);
                message.Subject = "Tách file hoàn tất";
                message.Body = "Hệ thống đã xử lý file Excel và tách các nhóm thành công. File đính kèm chứa tất cả các nhóm.";
                message.Attachments.Add(new Attachment(zipPath));

                using var smtp = new SmtpClient("smtp.gmail.com", 587)
                {
                    EnableSsl = true,
                    Credentials = new NetworkCredential(mail, appPassword)
                };

                await smtp.SendMailAsync(message);
                StatusMessage = $"Đã gửi email thành công tới {RecipientEmail}";
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("", "Gửi email thất bại: " + ex.Message);
            }
        }

        try
        {
            System.IO.File.Delete(zipPath);
        }
        catch
        {
            // ignore delete error
        }

        return Page();
    }

    public IActionResult OnPostDownloadAll()
    {
        var downloadFolder = Path.Combine("wwwroot", "downloads");
        var zipPath = Path.Combine(Path.GetTempPath(), $"AllFiles_{Guid.NewGuid()}.zip");

        var filesToZip = Directory.GetFiles(downloadFolder, "*.xlsx");

        using (var zipStream = new FileStream(zipPath, FileMode.Create))
        {
            using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var file in filesToZip)
                {
                    archive.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
        }

        var memory = new MemoryStream();
        using (var stream = new FileStream(zipPath, FileMode.Open, FileAccess.Read))
        {
            stream.CopyTo(memory);
        }
        memory.Position = 0;

        try
        {
            System.IO.File.Delete(zipPath);
        }
        catch
        {
            // ignore delete error
        }

        return File(memory, "application/zip", "AllFiles.zip");
    }
}