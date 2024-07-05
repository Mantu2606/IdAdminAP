/*using AutoMapper;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ILakshya.Dal;
using ILakshya.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenXmlCellType = DocumentFormat.OpenXml.Spreadsheet.CellType;
using NpoiCellType = NPOI.SS.UserModel.CellType;
using MathNet.Numerics.Distributions;
using NPOI.OpenXmlFormats.Dml.Diagram;
namespace ILakshya.WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class StudentController : ControllerBase
    {
        private readonly ICommonRepository<Student> _studentRepository;
        private readonly ICommonRepository<StudentMarks> _marksRepository;
        private readonly WebPocHubDbContext _dbContext;
        private readonly IMapper _mapper;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public StudentController(WebPocHubDbContext dbContext, ICommonRepository<Student> repository, IMapper mapper, IWebHostEnvironment webHostEnvironment, ICommonRepository<StudentMarks> marksRepository)
        {
            _dbContext = dbContext;
            _studentRepository = repository;
<<<<<<< HEAD
            _mapper = mapper;
            _webHostEnvironment = webHostEnvironment;
            _marksRepository = marksRepository;
=======
        }

        // Excel file upload endpoint
        [HttpPost("UploadExcel")]
         public async Task<IActionResult> UploadExcel(IFormFile file)
         {
             if (file == null || file.Length == 0)
             {
                 return BadRequest("No file uploaded.");
             }
             var students = new List<Student>();
             using (var stream = new MemoryStream())
             {
                 await file.CopyToAsync(stream);
                 stream.Position = 0; // Reset the stream position to the beginning
                 using (SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, false))
                 {
                     WorkbookPart workbookPart = doc.WorkbookPart;
                     Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault();
                     WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                     SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                     var headers = new List<string>();
                     bool isFirstRow = true;

                     var existingStudents = _dbContext.Students.ToDictionary(s => s.EnrollNo);
                     foreach (Row row in sheetData.Elements<Row>())
                     {
                         if (isFirstRow)
                         {
                             // Read headers
                             headers = row.Elements<Cell>().Select(cell => GetCellValue(doc, cell)).ToList();
                             isFirstRow = false;
                             continue;
                         }
                         var student = new Student();
                         var cells = row.Elements<Cell>().ToArray();
                         if (cells.Length < 14) // Ensure there are enough cells
                         {
                             continue; // Skip rows with insufficient data
                         }

                       //student.Id = GenerateStudentId(); // Explicitly set the Id
                         student.EnrollNo = cells.Length > 0 ? ParseCellValue(cells[0], doc) : null;

                         if (student.EnrollNo != null && existingStudents.TryGetValue(student.EnrollNo, out var existingStudent))
                         {
                             // If exists, update the existing student instead of adding a new one
                             student = existingStudent;
                         }

                         student.Name = cells.Length > 1 ? GetCellValue(doc, cells[1]) : "Unknown";
                         var fatherNameCellValue = cells.Length > 2 ? GetCellValue(doc, cells[2]) : null;
                         student.FatherName = cells.Length > 2 ? GetCellValue(doc, cells[2]) : "Unknown";

                         student.RollNo = cells.Length > 3 ? ParseCellValue(cells[3], doc).ToString() : null;
                         student.GenKnowledge = cells.Length > 4 ? ParseCellValue(cells[4], doc) ?? 0 : 0;
                         student.Science = cells.Length > 5 ? ParseCellValue(cells[5], doc) ?? 0 : 0;
                         student.EnglishI = cells.Length > 6 ? ParseCellValue(cells[6], doc) ?? 0 : 0;
                         student.EnglishII = cells.Length > 7 ? ParseCellValue(cells[7], doc) ?? 0 : 0;
                         student.HindiI = cells.Length > 8 ? ParseCellValue(cells[8], doc) ?? 0 : 0;
                         student.HindiII = cells.Length > 9 ? ParseCellValue(cells[9], doc) ?? 0 : 0;
                         student.Computer = cells.Length > 10 ? ParseCellValue(cells[10], doc) ?? 0 : 0;
                         student.Sanskrit = cells.Length > 11 ? ParseCellValue(cells[11], doc) ?? 0 : 0;
                         student.Mathematics = cells.Length > 12 ? ParseCellValue(cells[12], doc) ?? 0 : 0;
                         student.SocialStudies = cells.Length > 13 ? ParseCellValue(cells[13], doc) ?? 0 : 0;
                         student.MaxMarks = 5;  // Assuming max marks are 5 for all subjects
                         student.PassMarks = 2; // Assuming pass marks are 2 for all subjects

                         students.Add(student);
                     }
                 }
             }

             try
             {
                 _dbContext.Students.AddRange(students);
                 await _dbContext.SaveChangesAsync();
             }
             catch (Exception ex)
             {
                 return StatusCode(500, $"Internal server error: {ex.Message}");
             }

             return Ok(students); // Return the students
         }


        private int GenerateStudentId()
        {
            // You can implement your logic to generate a unique ID here
            // For example, you can query the database to get the last used ID and increment it by 1
            var lastStudent = _dbContext.Students.OrderByDescending(s => s.Id).FirstOrDefault();
            return (lastStudent != null ? lastStudent.Id : 0) + 1;
        }

        private int? ParseCellValue(Cell cell, SpreadsheetDocument doc)
        {
            string value = GetCellValue(doc, cell);
            if (value == null)
                return null;

            int parsedValue;
            if (int.TryParse(value, out parsedValue))
                return parsedValue;
            else
                return null; // Return null if parsing fails
        }

        private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            SharedStringTablePart sstPart = doc.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue?.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return sstPart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
            }
            return value;
>>>>>>> 8d8ab89446f17c4e4a25ad148ffbeec73fbb73eb
        }

        [HttpGet]
        public IEnumerable<Student> GetAll()
        {
            return _studentRepository.GetAll();
        }

        // This is work anytype file excel

        [HttpPost("UploadExcel")]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            var exam = Request.QueryString;

            var students = new List<Student>();
            var logins = new List<User>();

            try
            {
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    stream.Position = 0;

                    IWorkbook workbook = null;

                    if (file.FileName.EndsWith(".Xls")) // For Excel 97-2003 format (XLS)
                    {
                        workbook = new HSSFWorkbook(stream);
                    }
                    else if (file.FileName.EndsWith(".xlsx")) // For Excel 2007+ format (XLSX)
                    {
                        workbook = new XSSFWorkbook(stream);
                    }

                    else if (file.FileName.EndsWith(".xls")) // For Excel 2007+ format (XLSX)
                    {
                        workbook = new HSSFWorkbook(stream);
                    }

                    else
                    {
                        return BadRequest("Unsupported file format. Please upload a .xls or .xlsx file.");
                    }

                    if (workbook == null)
                    {
                        return BadRequest("Unsupported file");
                    }
                    //  var sheet = workbook.GetSheetAt(0) original
                    var sheet = workbook.GetSheetAt(0); // Assuming only one sheet


                    // var existingStudents = _dbContext.Students.ToDictionary(s => s.EnrollNo);
                    var existingStudents = _dbContext.Students.ToDictionary(s => s.Id);
                    for (int rowIdx = 1; rowIdx <= sheet.LastRowNum; rowIdx++) // Start from 1 to skip header
                    {
                        var row = sheet.GetRow(rowIdx);
                        if (row == null) continue; // Skip empty rows

                        var student = new Student();
                       
                        student.EnrollNo = ParseCellValue(row.GetCell(0));
                        //   For enroll dubalicaty
                       /* if (student.EnrollNo != null && existingStudents.TryGetValue(student.EnrollNo, out var existingStudent))
                        {
                            student = existingStudent;
                        }*/

                        student.Name = GetCellValue(row.GetCell(1));
                        student.FatherName = GetCellValue(row.GetCell(2));
                        student.RollNo = ParseCellValue(row.GetCell(3))?.ToString();
                    /*    student.GenKnowledge = ParseCellValue(row.GetCell(4)) ?? 0;
                        student.Science = ParseCellValue(row.GetCell(5)) ?? 0;
                        student.EnglishI = ParseCellValue(row.GetCell(6)) ?? 0;
                        student.EnglishII = ParseCellValue(row.GetCell(7)) ?? 0;
                        student.HindiI = ParseCellValue(row.GetCell(8)) ?? 0;
                        student.HindiII = ParseCellValue(row.GetCell(9)) ?? 0;
                        student.Computer = ParseCellValue(row.GetCell(10)) ?? 0;
                        student.Sanskrit = ParseCellValue(row.GetCell(11)) ?? 0;
                        student.Mathematics = ParseCellValue(row.GetCell(12)) ?? 0;
                        student.SocialStudies = ParseCellValue(row.GetCell(13)) ?? 0;
                        student.MaxMarks = ParseCellValue(row.GetCell(14)) ?? 0;  //= 5;// // Example values, adjust as needed
                        student.PassMarks = ParseCellValue(row.GetCell(15)) ?? 0; //= 2;
*/
                        if (student.RollNo == null)
                            continue;

                        students.Add(student);

                        logins.Add(new User()
                        {
                            Email = student.EnrollNo.ToString(),
                            EnrollNo = student.EnrollNo.ToString(),
                            Password = BCrypt.Net.BCrypt.HashPassword(student.EnrollNo + "_p@11"),
                            RoleId = 2
                        }
                        );
                    }
                }
                _dbContext.Students.AddRange(students);
                _dbContext.Users.AddRange(logins);
                await _dbContext.SaveChangesAsync();

                return Ok(students);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }



        [HttpPost("UploadMarksExcel")]
        public async Task<IActionResult> UploadMarksExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }
            var exam = Request.Query["exam"].ToString(); //Sachine sir
            var students = new List<StudentMarks>();


            try
            {
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    stream.Position = 0;

                    IWorkbook workbook = null;

                    if (file.FileName.EndsWith(".Xls")) // For Excel 97-2003 format (XLS)
                    {
                        workbook = new HSSFWorkbook(stream);
                    }
                    else if (file.FileName.EndsWith(".xlsx")) // For Excel 2007+ format (XLSX)
                    {
                        workbook = new XSSFWorkbook(stream);
                    }
                    else if (file.FileName.EndsWith(".xls")) // For Excel 2007+ format (XLSX)
                    {
                        workbook = new HSSFWorkbook(stream);
                    }

                    else
                    {
                        return BadRequest("Unsupported file format. Please upload a .xls or .xlsx file.");
                    }

                    if (workbook == null)
                    {
                        return BadRequest("Unsupported file");
                    }
                    //  var sheet = workbook.GetSheetAt(0) original
                    var sheet = workbook.GetSheetAt(0); // Assuming only one sheet

                    // var existingStudents = _dbContext.Students.ToDictionary(s => s.EnrollNo);
                  //  var existingStudents = _dbContext.Students.ToDictionary(s => s.Id);
                    for (int rowIdx = 1; rowIdx <= sheet.LastRowNum; rowIdx++) // Start from 1 to skip header
                    {
                        var row = sheet.GetRow(rowIdx);
                        if (row == null) continue; // Skip empty rows

                        var markdata = new StudentMarks();
                        markdata.Exam = exam;
                        markdata.EnrollNo = ParseCellValue(row.GetCell(0));
                        markdata.RollNo = ParseCellValue(row.GetCell(3))?.ToString();
                        markdata.GenKnowledge = ParseCellValue(row.GetCell(4)) ?? 0;
                        markdata.Science = ParseCellValue(row.GetCell(5)) ?? 0;
                        markdata.EnglishI = ParseCellValue(row.GetCell(6)) ?? 0;
                        markdata.EnglishII = ParseCellValue(row.GetCell(7)) ?? 0;
                        markdata.HindiI = ParseCellValue(row.GetCell(8)) ?? 0;
                        markdata.HindiII = ParseCellValue(row.GetCell(9)) ?? 0;
                        markdata.Computer = ParseCellValue(row.GetCell(10)) ?? 0;
                        markdata.Sanskrit = ParseCellValue(row.GetCell(11)) ?? 0;
                        markdata.Mathematics = ParseCellValue(row.GetCell(12)) ?? 0;
                        markdata.SocialStudies = ParseCellValue(row.GetCell(13)) ?? 0;


                        markdata.MaxMarks = ParseCellValue(row.GetCell(14)) ?? 0;  //= 5;/ // Example values, adjust as needed
                        markdata.PassMarks = ParseCellValue(row.GetCell(15)) ?? 0; //= 2;/

                        if (markdata.RollNo == null)
                            continue;

                        students.Add(markdata);
                    }
                }
                _dbContext.StudentMarks.AddRange(students);
                await _dbContext.SaveChangesAsync();

                return Ok(students);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }


        private string GetCellValue(NPOI.SS.UserModel.ICell cell)
        {
            if (cell == null) return null;

            switch (cell.CellType)
            {
                case NpoiCellType.String:
                    return cell.StringCellValue;
                case NpoiCellType.Numeric:
                    if (NPOI.SS.UserModel.DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue.ToString(); // Handle date values as needed
                    else
                        return cell.NumericCellValue.ToString();
                case NpoiCellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case NpoiCellType.Formula:
                    return cell.CellFormula; // Handle formula if needed
                default:
                    return null;
            }
        }
        private int? ParseCellValue(NPOI.SS.UserModel.ICell cell)
        {
            if (cell == null || cell.CellType == NpoiCellType.Blank)
                return null;

            switch (cell.CellType)
            {
                case NpoiCellType.Numeric:
                    return (int)Math.Round(cell.NumericCellValue);
                case NpoiCellType.String:
                    if (int.TryParse(cell.StringCellValue, out int intValue))
                        return intValue;
                    return null;
                default:
                    return null;
            }
        }

        [HttpGet("{id:int}")]
        public ActionResult<Student> GetById(int id)
        {
            var student = _studentRepository.GetDetails(id);
            if (student == null)
            {
                return NotFound();
            }
            return Ok(student);
        }

        [HttpGet("ByEnrollNo/{enrollNo}")]
        public ActionResult<Student> GetStudentDetailsByEnrollNo(string enrollNo)
        {
            if (string.IsNullOrEmpty(enrollNo))
            {
                return BadRequest("EnrollNo cannot be null or empty.");
            }
            
            var student = _studentRepository.GetAll()
                .FirstOrDefault(s => s.EnrollNo != null && s.EnrollNo.ToString() == enrollNo);
            student.studentmarks  = _marksRepository.GetAll().Where(x => x.EnrollNo == student.EnrollNo).ToList<StudentMarks>();

            if (student == null)
            {
                return NotFound("Student Not found");
            }

            return Ok(student);
        }

        [HttpPost("UploadProfilePicture/{id}")]
        public async Task<IActionResult> UploadProfilePicture(int id, IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            var student = _studentRepository.GetDetails(id);
            if (student == null)
            {
                return NotFound();
            }

            var uploadsFolder = Path.Combine(_webHostEnvironment.WebRootPath, "profile_pictures");
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }

            var uniqueFileName = $"{id}_{Path.GetRandomFileName()}{Path.GetExtension(file.FileName)}";
            var filePath = Path.Combine(uploadsFolder, uniqueFileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            student.ProfilePicture = $"profile_pictures/{uniqueFileName}";
            _studentRepository.Update(student);
            await _studentRepository.SaveChangesAsync();

            return Ok(new { student.ProfilePicture });
        }

        [HttpDelete("{id}")]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public ActionResult<Student> Delete(int id)
        {
            var student = _studentRepository.GetDetails(id);
            if (student == null) return NotFound();

            _studentRepository.Delete(student);
            _studentRepository.SaveChanges();
            return NoContent();
        }

        [HttpDelete("ByEnrollNo/{enrollNo}")]
        public ActionResult<Student> DeleteByEnrollNo(string enrollNo)
        {
            var student = _studentRepository.GetAll().FirstOrDefault(s => s.EnrollNo?.ToString() == enrollNo);
            if (student == null)
            {
                return NotFound();
            }

            _studentRepository.Delete(student);
            _studentRepository.SaveChanges();

            return NoContent();
        }
    }
}
<<<<<<< HEAD

=======
*/


using AutoMapper;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ILakshya.Dal;
using ILakshya.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ILakshya.WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class StudentController : ControllerBase
    {
        private readonly ICommonRepository<Student> _studentRepository;
        private readonly WebPocHubDbContext _dbContext;

        public StudentController(WebPocHubDbContext dbContext, ICommonRepository<Student> repository, IMapper mapper)
        {
            _dbContext = dbContext;
            _studentRepository = repository;
        }

        [HttpGet]
        public IEnumerable<Student> GetAll()
        {
            return _studentRepository.GetAll();
        }

        [HttpGet("{id:int}")]
        public ActionResult<Student> GetById(int id)
        {
            var student = _studentRepository.GetDetails(id);
            if (student == null)
            {
                return NotFound();
            }
            return Ok(student);
        }

        [HttpGet("ByEnrollNo/{enrollNo}")]
        public ActionResult<Student> GetStudentDetailsByEnrollNo(string enrollNo)
        {
            // Ensure enrollNo is not null or empty
            if (string.IsNullOrEmpty(enrollNo))
            {
                return BadRequest("EnrollNo cannot be null or empty.");
            }

            // Find the student by enrollNo
            var student = _studentRepository.GetAll().FirstOrDefault(s => s.EnrollNo != null && s.EnrollNo.ToString() == enrollNo);

            // If student is not found
            if (student == null)
            {
                return NotFound("Student Not found");
            }

            return Ok(student);
        }

        [HttpPost("UploadExcel")]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            var students = new List<Student>();

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                stream.Position = 0;

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault();
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    var headers = new List<string>();
                    bool isFirstRow = true;

                    var existingStudents = _dbContext.Students.ToDictionary(s => s.EnrollNo);
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        if (isFirstRow)
                        {
                            headers = row.Elements<Cell>().Select(cell => GetCellValue(doc, cell)).ToList();
                            isFirstRow = false;
                            continue;
                        }

                        var student = new Student();
                        var cells = row.Elements<Cell>().ToArray();
                        if (cells.Length < 14) continue;

                        student.EnrollNo = cells.Length > 0 ? ParseCellValue(cells[0], doc) : null;

                        if (student.EnrollNo != null && existingStudents.TryGetValue(student.EnrollNo, out var existingStudent))
                        {
                            student = existingStudent;
                        }

                        student.Name = cells.Length > 1 ? GetCellValue(doc, cells[1]) : "Unknown";
                        student.FatherName = cells.Length > 2 ? GetCellValue(doc, cells[2]) : "Unknown";
                        student.RollNo = cells.Length > 3 ? ParseCellValue(cells[3], doc).ToString() : null;
                        student.GenKnowledge = cells.Length > 4 ? ParseCellValue(cells[4], doc) ?? 0 : 0;
                        student.Science = cells.Length > 5 ? ParseCellValue(cells[5], doc) ?? 0 : 0;
                        student.EnglishI = cells.Length > 6 ? ParseCellValue(cells[6], doc) ?? 0 : 0;
                        student.EnglishII = cells.Length > 7 ? ParseCellValue(cells[7], doc) ?? 0 : 0;
                        student.HindiI = cells.Length > 8 ? ParseCellValue(cells[8], doc) ?? 0 : 0;
                        student.HindiII = cells.Length > 9 ? ParseCellValue(cells[9], doc) ?? 0 : 0;
                        student.Computer = cells.Length > 10 ? ParseCellValue(cells[10], doc) ?? 0 : 0;
                        student.Sanskrit = cells.Length > 11 ? ParseCellValue(cells[11], doc) ?? 0 : 0;
                        student.Mathematics = cells.Length > 12 ? ParseCellValue(cells[12], doc) ?? 0 : 0;
                        student.SocialStudies = cells.Length > 13 ? ParseCellValue(cells[13], doc) ?? 0 : 0;
                        student.MaxMarks = 5;
                        student.PassMarks = 2;

                        students.Add(student);
                    }
                }
            }

            try
            {
                _dbContext.Students.AddRange(students);
                await _dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }

            return Ok(students);
        }

        [HttpDelete("{id}")]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public ActionResult<Student> Delete(int id)
        {
            var student = _studentRepository.GetDetails(id);
            if (student == null) return NotFound();

            _studentRepository.Delete(student);
            _studentRepository.SaveChanges();
            return NoContent();
        }

        [HttpDelete("ByEnrollNo/{enrollNo}")]
        public ActionResult<Student> DeleteByEnrollNo(string enrollNo)
        {
            var student = _studentRepository.GetAll().FirstOrDefault(s => s.EnrollNo?.ToString() == enrollNo);
            if (student == null)
            {
                return NotFound();
            }

            _studentRepository.Delete(student);
            _studentRepository.SaveChanges();

            return NoContent();
        }

        [HttpPut]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public ActionResult Update(Student student)
        {
            _studentRepository.Update(student);
            var result = _studentRepository.SaveChanges();
            return result > 0 ? NoContent() : (ActionResult)BadRequest();
        }

        private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            SharedStringTablePart sstPart = doc.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue?.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return sstPart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
            }
            return value;
        }

        private int? ParseCellValue(Cell cell, SpreadsheetDocument doc)
        {
            string value = GetCellValue(doc, cell);
            if (value == null)
                return null;

            if (int.TryParse(value, out int parsedValue))
                return parsedValue;

            return null;
        }
    }
}
>>>>>>> 8d8ab89446f17c4e4a25ad148ffbeec73fbb73eb
