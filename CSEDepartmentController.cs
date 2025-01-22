using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PDFFORMATDATA.DataBase;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Pdf.Canvas;
using iText.Layout.Properties;
using iText.Kernel.Events;
using System.Globalization;
using OfficeOpenXml;
using OfficeOpenXml.Style;

[Route("api/[controller]")]
[ApiController]
public class CSEDepartmentController : ControllerBase
{
    private readonly AppDbContext _context;

    public CSEDepartmentController(AppDbContext context)
    {
        _context = context;
    }

    [HttpGet("download")]
    public async Task<IActionResult> DownloadRecordsAsPdf()
    {
        var records = await _context.CSEDepartments.ToListAsync();

        if (records == null || records.Count == 0)
        {
            return NotFound("No records found.");
        }

        using (var memoryStream = new MemoryStream())
        {
            var writer = new PdfWriter(memoryStream);
            var pdf = new PdfDocument(writer);
            var document = new Document(pdf);

            // Add event handler for header and footer
            pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, new HeaderFooterEventHandler());

            // Add title
            document.Add(new Paragraph("CSE Department Student List")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBold()
                .SetFontSize(18)
                .SetMarginBottom(20));

            // Create table for records
            var table = new Table(UnitValue.CreatePercentArray(new float[] { 1, 3, 3, 2, 1, 3 }))
                .UseAllAvailableWidth()
                .SetMarginTop(10);

            // Add table headers
            table.AddHeaderCell(new Cell().Add(new Paragraph("ID").SetBold()));
            table.AddHeaderCell(new Cell().Add(new Paragraph("Student Name").SetBold()));
            table.AddHeaderCell(new Cell().Add(new Paragraph("Enrollment No").SetBold()));
            table.AddHeaderCell(new Cell().Add(new Paragraph("Course").SetBold()));
            table.AddHeaderCell(new Cell().Add(new Paragraph("Year").SetBold()));
            table.AddHeaderCell(new Cell().Add(new Paragraph("Contact / Email").SetBold()));

            // Populate table with data
            foreach (var record in records)
            {
                table.AddCell(new Paragraph(record.studendId.ToString()));
                table.AddCell(new Paragraph(record.StudentName));
                table.AddCell(new Paragraph(record.EnrollmentNumber));
                table.AddCell(new Paragraph(record.Course));
                table.AddCell(new Paragraph(record.Year.ToString()));
                table.AddCell(new Paragraph($"{record.ContactNumber ?? "N/A"} / {record.Email ?? "N/A"}"));
            }

            document.Add(table);
            document.Close();

            var pdfBytes = memoryStream.ToArray();
            return File(pdfBytes, "application/pdf", "CSEDepartmentRecords.pdf");
        }
    }
    // excel format data 

    [HttpGet("download/excel")]
    public async Task<IActionResult> DownloadRecordsAsExcel()
    {
        var records = await _context.CSEDepartments.ToListAsync();
        if(records == null || records.Count == 0)
        {
            return NotFound("No records found");
        }

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("CSE Department");
            worksheet.Cells["A1"].Value = "ID";
            worksheet.Cells["B1"].Value = "Student Name";
            worksheet.Cells["C1"].Value = "Enrollment No";
            worksheet.Cells["D1"].Value = "Course";
            worksheet.Cells["E1"].Value = "Year";
            worksheet.Cells["F1"].Value = "Contact / Email";

            using (var range = worksheet.Cells["A1:F1"])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            int row = 2;
            foreach (var record in records)
            {
                worksheet.Cells[row, 1].Value = record.studendId;
                worksheet.Cells[row, 2].Value = record.StudentName;
                worksheet.Cells[row, 3].Value = record.EnrollmentNumber;
                worksheet.Cells[row, 4].Value = record.Course;
                worksheet.Cells[row, 5].Value = record.Year;
                worksheet.Cells[row, 6].Value = $"{record.ContactNumber ?? "N/A"} / {record.Email ?? "N/A"}";
                row++;
            }
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            var excelBytes = package.GetAsByteArray();
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CSEDepartmentRecords.xlsx");



        }

    }

    private class HeaderFooterEventHandler : IEventHandler
    {
        public void HandleEvent(Event @event)
        {
            var pdfEvent = (PdfDocumentEvent)@event;
            var page = pdfEvent.GetPage();
            var pdfCanvas = new PdfCanvas(page.NewContentStreamAfter(), page.GetResources(), pdfEvent.GetDocument());
            var pageSize = page.GetPageSize();
            var canvas = new Canvas(pdfCanvas, pageSize);

            // Get current date and time
            string currentDateTime = DateTime.Now.ToString("MMMM dd, yyyy hh:mm tt", CultureInfo.InvariantCulture);

            // Add current date and time to the header
            canvas
                .ShowTextAligned($"Generated on: {currentDateTime}",
                    pageSize.GetLeft() + 40, // Left margin
                    pageSize.GetTop() - 20,  // Top margin
                    TextAlignment.LEFT);

            // Add page number to the footer
            int pageNumber = pdfEvent.GetDocument().GetPageNumber(page);
            canvas
                .ShowTextAligned($"Page: {pageNumber}",
                    pageSize.GetRight() - 40, // Right margin
                    pageSize.GetBottom() + 20, // Bottom margin
                    TextAlignment.RIGHT);

            canvas.Close();
        }
    }
}
