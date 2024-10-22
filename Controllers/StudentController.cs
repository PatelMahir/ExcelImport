using ExcelImport.Data;
using ExcelImport.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
namespace ExcelImport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class StudentController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        public StudentController(ApplicationDbContext context)
        {
            _context = context;
        }

        #region Right POST Excel method
        [HttpPost("ImportExcel")]
        public async Task<IActionResult> ImportExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("File is empty");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);
            using var package = new ExcelPackage(stream);

            var worksheet = package.Workbook.Worksheets.First();
            var rowCount = worksheet.Dimension?.Rows ?? 0;  // Null check for Dimension

            for (int row = 2; row <= rowCount; row++)  // Assuming the first row is the header
            {
                var idValue = worksheet.Cells[row, 1]?.Value?.ToString()?.Trim();
                var nameValue = worksheet.Cells[row, 2]?.Value?.ToString()?.Trim();
                var marksValue = worksheet.Cells[row, 3]?.Value?.ToString()?.Trim();

                // Check if the entire row is empty (no data in any cell)
                if (string.IsNullOrWhiteSpace(idValue) && string.IsNullOrWhiteSpace(nameValue) && string.IsNullOrWhiteSpace(marksValue))
                {
                    continue;  // Skip this empty row
                }

                if (string.IsNullOrWhiteSpace(idValue) || string.IsNullOrWhiteSpace(nameValue) || string.IsNullOrWhiteSpace(marksValue))
                {
                    return BadRequest($"Invalid data at row {row}");
                }

                if (!int.TryParse(idValue, out var id) || !int.TryParse(marksValue, out var marks))
                {
                    return BadRequest($"Invalid data at row {row}: ID or Marks are not valid integers.");
                }

                var student = await _context.students.FindAsync(id);

                if (student != null)
                {
                    // Updating existing student record, do not modify Id.
                    student.Name = nameValue;
                    student.Marks = marks;
                }
                else
                {
                    // Creating a new student, do not set the Id.
                    student = new Student
                    {
                        Name = nameValue,
                        Marks = marks
                        // Do NOT set the Id here.
                    };
                    _context.students.Add(student);
                }
            }

            await _context.SaveChangesAsync();
            return Ok("Data imported successfully");
        }
        #endregion
        #region Export In form of Excel
        [HttpGet("export-excel")]
        public IActionResult ExportDataToExcel()
        {
            var employees = _context.students.ToList(); // Fetch data from DB
            
            // Set the license context for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  // Add this line

            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("students");

                // Add headers
                worksheet.Cells[1, 1].Value = "Id";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 3].Value = "Marks";

                var students = _context.students.ToList();
                if (students == null || !students.Any())
                {
                    return BadRequest("No student data available.");
                }
                // Add data from database
                int row = 2; // Start from row 2 (row 1 has headers)
                foreach (var student in students)
                {
                    worksheet.Cells[row, 1].Value = student.Id;
                    worksheet.Cells[row, 2].Value = student.Name;
                    worksheet.Cells[row, 3].Value = student.Marks;
                    row++;
                }

                // Auto-fit columns for better formatting
                worksheet.Cells.AutoFitColumns();

                // Return the Excel file as a downloadable response
                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;
                string excelName = $"Students-{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
            }
        }
        #endregion

    }
}
#region HttpPost
//[HttpPost("UploadExcelFile")]
//public IActionResult UploadExcelFile(IFormFile file)
////public IActionResult UploadExcelFile([FromBody] IFormFile file)
////public IActionResult UploadExcelFile([FromForm] string name, IFormFile file)
//{
//    if (file == null || file.Length == 0)
//    {
//        return BadRequest("No file uploaded or the file is empty.");
//    }

//    try
//    {
//        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
//        if (file == null || file.Length == 0)
//        {
//            return BadRequest("No File Uploaded");
//        }
//        var uploadsFolder = $"{Directory.GetCurrentDirectory()}\\Uploads";

//        if (!Directory.Exists(uploadsFolder))
//        {
//            Directory.CreateDirectory(uploadsFolder);
//        }

//        var filePath = Path.Combine(uploadsFolder, file.Name);
//        using (var stream = new FileStream(filePath, FileMode.Create))
//        {
//            file.CopyTo(stream);
//        }

//        using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
//        {
//            using (var reader = ExcelReaderFactory.CreateReader(stream))
//            {
//                bool isHeaderSkipped = false;
//                do
//                {
//                    while (reader.Read())
//                    {
//                        if (!isHeaderSkipped)
//                        {
//                            isHeaderSkipped = true;
//                            continue;
//                        }
//                        //reader.GetDouble(0);
//                        Student s = new Student();
//                        s.Name = reader.GetValue(1).ToString();
//                        s.Marks = Convert.ToInt32(reader.GetValue(2));

//                        _context.Add(s);
//                        _context.SaveChangesAsync();
//                    }
//                } while (reader.NextResult());
//            }
//        }

//        return Ok("Successfully Inserted");
//    }
//    catch (Exception ex)
//    {
//        return StatusCode(500, ex.Message);
//    }
//}

//[HttpPost("UploadExcelFile")]
//public IActionResult UploadExcelFile(IFormFile file)

////public IActionResult UploadExcelFile([FromForm] string name, IFormFile file)
//{
//    try
//    {
//        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

//        if (file == null || file.Length == 0)
//        {
//            return BadRequest("No File Uploaded");
//        }

//        var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Uploads");

//        if (!Directory.Exists(uploadsFolder))
//        {
//            Directory.CreateDirectory(uploadsFolder);
//        }

//        var fileName = Path.GetFileName(file.FileName);
//        var filePath = Path.Combine(uploadsFolder, fileName);

//        using (var stream = new FileStream(filePath, FileMode.Create))
//        {
//            file.CopyTo(stream);
//        }

//        using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
//        {
//            using (var reader = ExcelReaderFactory.CreateReader(stream))
//            {
//                bool isHeaderSkipped = false;
//                using (var transaction = _context.Database.BeginTransaction())
//                {
//                    try
//                    {
//                        do
//                        {
//                            while (reader.Read())
//                            {
//                                if (!isHeaderSkipped)
//                                {
//                                    isHeaderSkipped = true;
//                                    continue;
//                                }

//                                Student s = new Student
//                                {
//                                    Name = reader.GetValue(1)?.ToString(),
//                                    Marks = int.TryParse(reader.GetValue(2)?.ToString(), out var marks) ? marks : 0
//                                };

//                                _context.Add(s);
//                            }
//                        } while (reader.NextResult());

//                        _context.SaveChanges();
//                        transaction.Commit();
//                    }
//                    catch
//                    {
//                        transaction.Rollback();
//                        throw;
//                    }
//                }
//            }
//        }

//        return Ok("Successfully Inserted");
//    }
//    catch (Exception ex)
//    {
//        return StatusCode(500, $"Internal server error: {ex.Message}");
//    }
//}
#endregion


//[HttpPost("UploadExcelFile")]
//public async Task<IActionResult> UploadExcelFile(IFormFile file)
//{
//    if (file == null || file.Length == 0)
//    {
//        return BadRequest("No file uploaded or the file is empty.");
//    }

//    try
//    {
//        // Register encoding provider
//        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

//        // Create the uploads folder if it doesn't exist
//        var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Uploads");
//        if (!Directory.Exists(uploadsFolder))
//        {
//            Directory.CreateDirectory(uploadsFolder);
//        }

//        // Save the uploaded file to the uploads folder
//        var filePath = Path.Combine(uploadsFolder, file.FileName);
//        using (var stream = new FileStream(filePath, FileMode.Create))
//        {
//            await file.CopyToAsync(stream);
//        }

//        // Open the file for reading
//        using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
//        {
//            using (var reader = ExcelReaderFactory.CreateReader(stream))
//            {
//                bool isHeaderSkipped = false;

//                do
//                {
//                    while (reader.Read())
//                    {
//                        if (!isHeaderSkipped)
//                        {
//                            isHeaderSkipped = true;
//                            continue;  // Skip the header row
//                        }

//                        // Create a new student object
//                        Student s = new Student();

//                        //if(!reader.IsDBNull(0) && int.TryParse(reader.GetValue(0)?.ToString(), out int id))
//                        //{
//                        //    s.Id = id;
//                        //}
//                        // Check if the Name column (index 01) is null
//                        if (!reader.IsDBNull(1))
//                        {
//                            s.Name = reader.GetValue(1).ToString();
//                        }
//                        else
//                        {
//                            break;
//                            return BadRequest("Student Name cannot be null.");
//                        }

//                        // Check if the Marks column (index 2) is null and validate the data
//                        if (!reader.IsDBNull(2) && int.TryParse(reader.GetValue(2)?.ToString(), out int marks))
//                        {
//                            s.Marks = marks;
//                        }
//                        else
//                        {
//                            return BadRequest("Invalid Marks value. It should be a valid number.");
//                        }

//                        // Add the student to the database
//                        _context.Add(s);
//                    }

//                    // Save all the changes after processing the entire file
//                    await _context.SaveChangesAsync();

//                } while (reader.NextResult());
//            }
//        }

//        return Ok("Data successfully inserted into the database.");
//    }
//    catch (Exception ex)
//    {
//        // Return a detailed error message to help with debugging
//        return StatusCode(500, $"An error occurred: {ex.Message}");
//    }
//}

//    #region HttpUploadFile
//        [HttpPost("UploadExcelFile")]
//        public async Task<IActionResult> UploadExcelFile(IFormFile file)
//        {
//            if (file == null || file.Length == 0)
//            {
//                return BadRequest("No file uploaded or the file is empty.");
//            }

//            try
//            {
//                // Register encoding provider
//                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

//                // Create the uploads folder if it doesn't exist
//                var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Uploads");
//                if (!Directory.Exists(uploadsFolder))
//                {
//                    Directory.CreateDirectory(uploadsFolder);
//                }

//                // Save the uploaded file to the uploads folder
//                var filePath = Path.Combine(uploadsFolder, file.FileName);
//                using (var stream = new FileStream(filePath, FileMode.Create))
//                {
//                    await file.CopyToAsync(stream);
//                }

//                // Open the file for reading
//                using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
//                {
//                    using (var reader = ExcelReaderFactory.CreateReader(stream))
//                    {
//                        bool isHeaderSkipped = false;

//                        do
//                        {
//                            while (reader.Read())
//                            {
//                                if (!isHeaderSkipped)
//                                {
//                                    isHeaderSkipped = true;
//                                    continue;  // Skip the header row
//                                }

//                                // Create a new student object
//                                Student s = new Student();

//                                //if(!reader.IsDBNull(0) && int.TryParse(reader.GetValue(0)?.ToString(), out int id))
//                                //{
//                                //    s.Id = id;
//                                //}
//                                // Check if the Name column (index 01) is null
//                                if (!reader.IsDBNull(1))
//                                {
//                                    s.Name = reader.GetValue(1).ToString();
//                                }
//                                else
//                                {
//                                    break;
//                                    return BadRequest("Student Name cannot be null.");
//                                }

//                                // Check if the Marks column (index 2) is null and validate the data
//                                if (!reader.IsDBNull(2) && int.TryParse(reader.GetValue(2)?.ToString(), out int marks))
//                                {
//                                    s.Marks = marks;
//                                }
//                                else
//                                {
//                                    return BadRequest("Invalid Marks value. It should be a valid number.");
//                                }

//                                // Add the student to the database
//                                _context.Add(s);
//                            }

//                            // Save all the changes after processing the entire file
//                            await _context.SaveChangesAsync();

//                        } while (reader.NextResult());
//                    }
//                }

//                return Ok("Data successfully inserted into the database.");
//            }
//            catch (Exception ex)
//            {
//                // Return a detailed error message to help with debugging
//                return StatusCode(500, $"An error occurred: {ex.Message}");
//            }
//        }


//    }

//    #endregion

//    }
//}


#region Import Excel
//namespace ExcelImport.Controllers
//{
//    [Route("api/[controller]")]
//    [ApiController]
//    public class FileUploadController : ControllerBase
//    {
//        [HttpPost("upload-excel")]
//        public async Task<IActionResult> UploadExcelFile(IFormFile file)
//        {
//            if (file == null || file.Length == 0)
//            {
//                return BadRequest("No file uploaded or file is empty.");
//            }

//            try
//            {
//                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
//                using (var stream = new MemoryStream())
//                {
//                    await file.CopyToAsync(stream)
//;
//                    stream.Position = 0;

//                    using (var reader = ExcelReaderFactory.CreateReader(stream))
//                    {
//                        var result = reader.AsDataSet();
//                        var dataTable = result.Tables[0];

//                        // Read the data from the DataTable
//                        foreach (DataRow row in dataTable.Rows)
//                        {
//                            // Example: Read first two columns
//                            var firstColumn = row[0]?.ToString();
//                            var secondColumn = row[1]?.ToString();

//                            // Process your data here (e.g., store in DB or log)
//                            // For demonstration, we'll just log it
//                            System.Diagnostics.Debug.WriteLine($"Column 1: {firstColumn}, Column 2: {secondColumn}");
//                        }

//                        // After processing all rows, throw an error
//                        throw new InvalidOperationException("Data is readed");
//                    }
//                }
//            }
//            catch (InvalidOperationException ex)
//            {
//                return BadRequest(ex.Message);
//            }
//            catch (Exception ex)
//            {
//                return StatusCode(500, $"Internal server error: {ex.Message}");
//            }
//        }
//    }

//}
#endregion