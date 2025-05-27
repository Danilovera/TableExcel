using Microsoft.AspNetCore.Mvc;
using WebApplication1.Interfaces;

namespace WebApplication1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController(IExcelService excelService) : ControllerBase
    {
        private readonly IExcelService _excelService = excelService;

        [HttpGet(Name = "DownloadExcel")]
        public async Task<IActionResult> Get()
        {
            var content =  await _excelService.ReturnExcelFile();

            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Bitacora Diaria"+ DateTime.Now.ToString()+".xlsx");
        }
    }
}
