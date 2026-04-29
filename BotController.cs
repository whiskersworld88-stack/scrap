// Controllers/BotController.cs
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using BPJSScrapper.Services;
using System.IO;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace BPJSScraper.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class BotController : ControllerBase
    {
        private static LasikBotService _botService = new LasikBotService();
        private static string _currentFilePath = null;
        private static int _totalRows = 0;
        private static int _currentRow = 0;
        private static string _lastLog = "";
        private static string _lastLogType = "info";
        private static bool _isFinished = false;

        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcel(IFormFile excelFile)
        {
            try
            {
                if (excelFile == null || excelFile.Length == 0)
                    return BadRequest(new { success = false, message = "File tidak ditemukan" });

                var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Uploads");
                if (!Directory.Exists(uploadsFolder))
                    Directory.CreateDirectory(uploadsFolder);

                var filePath = Path.Combine(uploadsFolder, excelFile.FileName);
                
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await excelFile.CopyToAsync(stream);
                }

                // Read total rows
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage(new FileInfo(filePath));
                var worksheet = package.Workbook.Worksheets[0];
                _totalRows = worksheet.Dimension?.End.Row - 1 ?? 0;

                _currentFilePath = filePath;
                _botService.ExcelPath = filePath;

                return Ok(new { 
                    success = true, 
                    message = "File berhasil diupload",
                    filePath = filePath,
                    totalRows = _totalRows
                });
            }
            catch (Exception ex)
            {
                return BadRequest(new { success = false, message = ex.Message });
            }
        }

        [HttpPost("open-chrome")]
        public IActionResult OpenChrome()
        {
            try
            {
                _botService.OpenChromeOnly();
                return Ok(new { success = true, message = "Chrome berhasil dibuka" });
            }
            catch (Exception ex)
            {
                return BadRequest(new { success = false, message = ex.Message });
            }
        }

        [HttpPost("start-bot")]
        public IActionResult StartBot()
        {
            try
            {
                if (string.IsNullOrEmpty(_currentFilePath))
                    return BadRequest(new { success = false, message = "Upload file Excel terlebih dahulu" });

                // Reset state
                _currentRow = 0;
                _isFinished = false;
                
                // Attach event handlers
                _botService.OnProgress = (current, total) => 
                {
                    _currentRow = current;
                };
                
                _botService.OnStatus = (message, isSuccess) => 
                {
                    _lastLog = message;
                    _lastLogType = isSuccess ? "success" : "info";
                    if (!isSuccess) _lastLogType = "error";
                };

                // Run bot in background
                Task.Run(() => 
                {
                    try
                    {
                        _botService.AttachBot();
                        _botService.Start();
                        _isFinished = true;
                        _lastLog = "✅ Bot selesai menjalankan semua data";
                        _lastLogType = "success";
                    }
                    catch (Exception ex)
                    {
                        _lastLog = $"❌ Error: {ex.Message}";
                        _lastLogType = "error";
                        _isFinished = true;
                    }
                });

                return Ok(new { success = true, message = "Bot dimulai" });
            }
            catch (Exception ex)
            {
                return BadRequest(new { success = false, message = ex.Message });
            }
        }

        [HttpPost("stop-bot")]
        public IActionResult StopBot()
        {
            try
            {
                _botService.Stop();
                return Ok(new { success = true, message = "Bot dihentikan" });
            }
            catch (Exception ex)
            {
                return BadRequest(new { success = false, message = ex.Message });
            }
        }

        [HttpGet("status")]
        public IActionResult GetStatus()
        {
            return Ok(new 
            { 
                success = true,
                currentRow = _currentRow,
                totalRows = _totalRows,
                successCount = _botService.SuccessCount,
                failedCount = _botService.FailedCount,
                lastLog = _lastLog,
                lastLogType = _lastLogType,
                isFinished = _isFinished,
                outputFile = _botService.ExcelPath?.Replace(".xlsx", "_HASIL.xlsx") ?? ""
            });
        }
    }
}
