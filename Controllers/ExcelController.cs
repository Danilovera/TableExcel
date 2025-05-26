using Microsoft.AspNetCore.Mvc;

namespace WebApplication1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpGet(Name = "DownloadExcel")]
        public async Task<IActionResult> Get()
        {
            await Task.Delay(1);
            var content = new byte[2024];

            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Bitacora Diaria"+ DateTime.Now.ToString()+".xlsx");
        }
    }
}
