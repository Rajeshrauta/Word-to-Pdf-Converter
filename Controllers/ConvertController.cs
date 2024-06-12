using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using SautinSoft.Document;
using Section = SautinSoft.Document.Section;

namespace WordToPdfConvertor.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ConvertController : ControllerBase
    {
        [HttpPost("word-to-pdf")]
        public async Task<IActionResult>ConvertDocxToPdf(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            try
            {
                string originalFileName = Path.GetFileNameWithoutExtension(file.FileName);
                if (string.IsNullOrEmpty(originalFileName))
                    originalFileName = "converted";

                // Load the DOCX file into the DocumentCore
                DocumentCore dc;
                using (var ms = new MemoryStream())
                {
                    await file.CopyToAsync(ms);
                    ms.Position = 0;
                    dc = DocumentCore.Load(ms, new DocxLoadOptions());
                }

                Section blankSection = new Section(dc);
                Paragraph blankParagraph = new Paragraph(dc);
                blankSection.Blocks.Add(blankParagraph);
                dc.Sections.Add(blankSection);


                byte[] pdfBytes;

                // Convert the document to PDF
                using (var msPdf = new MemoryStream())
                {
                    dc.Save(msPdf, new PdfSaveOptions());
                    //msPdf.Position = 0;
                    //return File(msPdf.ToArray(), "application/pdf", "converted.pdf");
                    pdfBytes = msPdf.ToArray();
                }

                // Load the PDF into PdfDocument (PdfSharp)
                using (var inputPdfStream = new MemoryStream(pdfBytes))
                using (var outputPdfStream = new MemoryStream())
                {
                    PdfDocument pdfDocument = PdfReader.Open(inputPdfStream, PdfDocumentOpenMode.Modify);

                    // Remove the last page
                    if (pdfDocument.PageCount > 0)
                    {
                        pdfDocument.Pages.RemoveAt(pdfDocument.PageCount - 1);
                    }

                    // Save the modified PDF to a MemoryStream
                    pdfDocument.Save(outputPdfStream);
                    outputPdfStream.Position = 0;

                    return File(outputPdfStream.ToArray(), "application/pdf", $"{originalFileName}.pdf");

                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}
