using ExcelDataReader;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc; 

namespace ImportExcel.Controllers
{ 
    public class HomeController : Controller
    {
        private static int _addedRows = 0;
        private static string _message = "";
        private static bool _success = false;   

        [HttpPost]
        public async Task<IActionResult> AddCustomersByExcel(IFormFile excel)
        {
            var filePath = "";
            _addedRows = 0;
            _message = "";
            try
            {
                var mainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "excel");
                var exists = Directory.Exists(mainPath);
                if (!exists)
                    Directory.CreateDirectory(mainPath);
                filePath = Path.Combine(mainPath, excel.FileName);
                await using (Stream stream = new FileStream(filePath, FileMode.Create))
                {
                    await excel.CopyToAsync(stream);
                }

                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                await using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    using var reader = ExcelReaderFactory.CreateReader(stream);

                    int header = 1, updateCounter = 0, counter = 0;
                    var needBreak = false;
                    do
                    {
                        if (needBreak)
                        {
                            break;
                        }

                        while (reader.Read()) //Each ROW
                        {
                            if (header == 1)
                            {
                                header = 0;
                                continue;
                            } 

                            //read first col
                            var code = reader.GetValue(0);
                            if (code == null)
                            {
                                needBreak = true;
                                break;
                            }

                            _addedRows++;

                            //read another row 

                        }
                    } while (reader.NextResult()); //Move to NEXT SHEET

                    _message = "عملیات موفق";
                    _success = true;
                    FileUploader.FileUploader.Delete(filePath); 
                }
            }
            catch (Exception e)
            {
                FileUploader.FileUploader.Delete(filePath);
                _message = $"داده های خط {_addedRows} چک شود";
                _success = false;
            }

            return Json(new
            {
                result = ""
            });
        }

        public async Task<IActionResult> ExcelStatus()
        {
            return Json(new
            {
                row = _addedRows,
                message = _message,
                success = _success,
            });
        }
    }
}

