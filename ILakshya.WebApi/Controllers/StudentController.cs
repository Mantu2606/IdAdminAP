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

        // Get student details by enrollNo

        // Get student details by enrollNo
        [HttpGet("{enrollNo}")]
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



        [HttpDelete("{id}")]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public ActionResult<Student> Delete(int id)
        {
            var student = _studentRepository.GetDetails(id);
            if (student == null) return NotFound();
            if (student.EnrollNo?.ToString() == id.ToString()) // Convert EnrollNo to string before comparison
            {
                _studentRepository.Delete(student);
                _studentRepository.SaveChanges();
                return NoContent();
            }
            else
            {
                return NotFound();
            }
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

    }
}
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
