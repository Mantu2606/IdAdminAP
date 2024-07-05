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
            _mapper = mapper;
            _webHostEnvironment = webHostEnvironment;
            _marksRepository = marksRepository;
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

