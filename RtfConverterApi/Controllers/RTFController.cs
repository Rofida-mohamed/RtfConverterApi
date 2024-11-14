using Microsoft.AspNetCore.Mvc;
using SpireDoc = Spire.Doc;  // Alias for Spire.Doc.Document
using System.IO;

namespace RtfConverterApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RTFController : ControllerBase
    {

        // POST: api/RTF/ConvertRtfToDocx
        [HttpPost("ConvertRtfToDocx")]
        public async Task<IActionResult> ConvertRtfToDocx(IFormFile rtfFile)
        {
            if (rtfFile == null || rtfFile.Length == 0)
            {
                return BadRequest("No RTF file uploaded.");
            }

            try
            {
                // Generate a temporary file path for the uploaded RTF file
                string tempRtfFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".rtf");

                // Save the uploaded RTF file to the temporary path
                using (var fileStream = new FileStream(tempRtfFilePath, FileMode.Create))
                {
                    await rtfFile.CopyToAsync(fileStream);
                }

                // Convert the RTF file to DOCX using Spire.Doc
                string tempDocxPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
                Spire.Doc.Document spireDoc = new Spire.Doc.Document();
                spireDoc.LoadFromFile(tempRtfFilePath, Spire.Doc.FileFormat.Rtf);
                spireDoc.SaveToFile(tempDocxPath, Spire.Doc.FileFormat.Docx);

                // Return the DOCX file for download
                var docxBytes = System.IO.File.ReadAllBytes(tempDocxPath);
                return File(docxBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "converted-file.docx");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
        // POST: api/RTF/ConvertRtfToDocx
        [HttpPost("ConvertRtfToDocxUrl")]
        public async Task<IActionResult> ConvertRtfToDocxURl(IFormFile rtfFile)
        {
            if (rtfFile == null || rtfFile.Length == 0)
            {
                return BadRequest("No RTF file uploaded.");
            }

            try
            {
                // Generate a temporary file path for the uploaded RTF file
                string tempRtfFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".rtf");

                // Save the uploaded RTF file to the temporary path
                using (var fileStream = new FileStream(tempRtfFilePath, FileMode.Create))
                {
                    await rtfFile.CopyToAsync(fileStream);
                }

                // Convert the RTF file to DOCX using Spire.Doc
                string tempDocxPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
                SpireDoc.Document spireDoc = new SpireDoc.Document();
                spireDoc.LoadFromFile(tempRtfFilePath, SpireDoc.FileFormat.Rtf);
                spireDoc.SaveToFile(tempDocxPath, SpireDoc.FileFormat.Docx);

                // Save the DOCX file to wwwroot/files folder
                string wwwrootPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files");

                // Ensure the directory exists
                if (!Directory.Exists(wwwrootPath))
                {
                    Directory.CreateDirectory(wwwrootPath);
                }

                // Define the public file path
                string publicFilePath = Path.Combine(wwwrootPath, Path.GetFileName(tempDocxPath));

                // Copy the DOCX file to the wwwroot/files directory
                System.IO.File.Copy(tempDocxPath, publicFilePath);

                // Create the URL to access the file
                string fileUrl = $"{Request.Scheme}://{Request.Host}/files/{Path.GetFileName(tempDocxPath)}";

                // Return the URL where the file can be accessed
                return Ok(new { fileUrl = fileUrl });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}
