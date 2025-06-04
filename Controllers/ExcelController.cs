using Microsoft.AspNetCore.Mvc;
using WebApplication1.Interfaces;
using WebApplication1.Models;

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
            var content =  await _excelService.ReturnExcelFile(GetDailyLogs(), GetProcedures(), "020");

            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Bitacora Diaria"+ DateTime.Now.ToString()+".xlsx");
        }

        private List<DailyLog> GetDailyLogs()
        {
            return new List<DailyLog>()
            {
                        //TODO: Crear los que faltan
        new DailyLog("343712","0104AUT0258111-01","AUTOS",new DateTime(2025,02,19),new DateTime(2025,03,19),63140,"V-508","RENOVACION","BELLIDO SOLANO LUIS JOSE","402440256",DateTime.Now,"50661393498","CL289812"),
        new DailyLog("343890","0104AUT0262031-00","AUTOS",new DateTime(2025,02,28),new DateTime(2025,03,30),32793,"V-515","RENOVACION","CHAVARRIA MORA JORGE STEVEN","113480170",DateTime.Now,"50688681066","BHL565","30"),
        new DailyLog("343770","0104AUT0170751-20","AUTOS",new DateTime(2025,02,21),new DateTime(2025,03,21),19615,"V-516","RENOVACION","BOLAÑOS CHAVES ANGIE TATIANA","111300987",DateTime.Now,"50662520001","BGP915","30"),
        new DailyLog("227831","0104AUT0264122-00","AUTOS",new DateTime(2025,02,28),new DateTime(2025,05,29),191552,"V-517","RENOVACION","BALMACEDA ARAGON ALEJANDRO","109760188",DateTime.Now,"50688587203","PLAN FAM VIQUEZ CRUZ MARCO VINICIO BDG307"),
        new DailyLog("346141","0104AUT0211279-13","AUTOS",new DateTime(2025,02,14),new DateTime(2025,03,14),21185,"V-519","RENOVACION","VIQUEZ BARQUERO YAMILETH","401390032",DateTime.Now,"50683635256","401753"),
        new DailyLog("292133","0104AUT0254422-02","AUTOS",new DateTime(2025,02,15),new DateTime(2025,03,15),35432,"V-520","RENOVACION","RAMOS GONZALEZ MELISSA","114020771",DateTime.Now,"50683303503","PLAN FAM ROJAS HIDALGO PABLO MRG021"),
        new DailyLog("292143","0104AUT0257398-01","AUTOS",new DateTime(2025,02,15),new DateTime(2025,03,15),22259,"V-520","RENOVACION","ROJAS ALVAREZ RODOLFO","105250453",DateTime.Now,"50683303503","PLAN FAM ROJAS HIDALGO PABLO BVP240"),
        new DailyLog("292158","0104AUT0242919-06","AUTOS",new DateTime(2025,02,15),new DateTime(2025,03,15),67099,"V-520","RENOVACION","ROJAS HIDALGO PABLO","111600202",DateTime.Now,"50683303503","PLAN FAM ROJAS HIDALGO PABLO GGG825"),
        new DailyLog("256403","0104AUT0255152-02","AUTOS",new DateTime(2025,02,15),new DateTime(2025,05,15),24400,"S-520","RENOVACION","ROJAS HIDALGO PABLO","111600202",DateTime.Now,"50683303503","74000"),
        new DailyLog("488596","0104AUT0265129-00","AUTOS",new DateTime(2025,02,15),new DateTime(2025,03,15),25320,"V-520","RENOVACION","ROJAS HIDALGO PABLO","111600202",DateTime.Now,"50683303503","PLAN FAM ROJAS HIDALGO PABLO AAG575"),
        new DailyLog("346039","0104AUT0251890-03","AUTOS",new DateTime(2025,02,16),new DateTime(2025,03,16),8665,"V-521","RENOVACION","SMITH SAENZ JADELYN SHADAY","402620232",DateTime.Now,"50685861010","875298"),
        new DailyLog("227751","0104AUT0257544-01","AUTOS",new DateTime(2025,02,14),new DateTime(2025,05,14),51742,"V-522","RENOVACION","ARIAS TORRENTES ANDREY DAVID","603470399",DateTime.Now,"50683028494","PLAN FAM ARIAS TORRENTES ANDREY DAVID BSM604"),
        new DailyLog("320210","0104AUT0266498-00","AUTOS",new DateTime(2025,02,16),new DateTime(2025,03,16),31979,"V-523","RENOVACION","GONZALEZ PEREZ JORGE ENRIQUE","602000204",DateTime.Now,"50688527020","CL313750"),
        new DailyLog("376302","0104AUT0267210-00","AUTOS",new DateTime(2025,02,16),new DateTime(2025,03,16),41526,"V-523","RENOVACION","GONZALEZ PEREZ JORGE ENRIQUE","602000204",DateTime.Now,"50688527020","CL349838"),
        new DailyLog("345359","0104AUT0264241-00","AUTOS",new DateTime(2025,02,28),new DateTime(2025,03,31),32403,"S-968620","RENOVACION","MADRIGAL LIZANO MARIA CELINA","108440638",DateTime.Now,"50687046532","PLAN FAM MADRIGAL LIZANO MARIA CELINA LST375"),
        new DailyLog("346101","0104AUT0249724-04","AUTOS",new DateTime(2025,02,18),new DateTime(2025,03,18),22220,"T-10433778","RENOVACION","TREJOS MENDEZ CAROLINA","116450108",DateTime.Now,"50684271114","802538"),
        new DailyLog("292040","0104AUT0256985-01","AUTOS",new DateTime(2025,02,14),new DateTime(2025,03,14),66023,"T-11271029","RENOVACION","MONGE ALFARO ALLAN","401580701",DateTime.Now,"50689807948","PLAN FAM MONGE ALFARO ALLAN RRS421"),
        new DailyLog("395161","0104IMR0007746-00","INCENDIO",new DateTime(2025,08,03),new DateTime(2025,08,04),13185,"E-EFECTIVO","RENOVACION","GUZMAN SOLANO DIRSEO GERARDO","111580425",DateTime.Now,"50685830037"),
        new DailyLog("400997","0104AUT0221549-12","AUTOS",new DateTime(2025,08,03),new DateTime(2025,08,04),14149,"E-EFECTIVO","RENOVACION","MERCADO GONZALEZ BYRON ANTONIO","155808624707",DateTime.Now,"50687128320","BQQ382"),
        new DailyLog("396475","0104IMR0006490-01","HOGAR COMPRENSIVO ",new DateTime(2025,09,03),new DateTime(2026,09,03),67949,"V-523","RENOVACION","MORA TORRES KATHERINE","114080036",DateTime.Now,"50661242278"),
        new DailyLog("488233","0104AUT0267097-00","AUTOS",new DateTime(2025,03,03),new DateTime(2025,03,06),140286,"V-518","RENOVACION","3 101 882035 S A","3101882035",DateTime.Now,"50670339600","DSY069"),
        new DailyLog("345134","0104AUT0261820-00","AUTOS",new DateTime(2025,02,16),new DateTime(2025,03,16),27090,"V-524","RENOVACION","HERRERA JIMENEZ JOHNNATAN ALEJANDRO","401930721",DateTime.Now,"50688095858","PMK090"),
        new DailyLog("345165","0104AUT0261801-00","AUTOS",new DateTime(2025,02,16),new DateTime(2025,03,16),20534,"V-524","RENOVACION","HERRERA JIMENEZ JOHNNATAN ALEJANDRO","401930721",DateTime.Now,"50688095858","BYD260"),
        new DailyLog("5612974"," ","AUTOS",new DateTime(2025,02,28),new DateTime(2025,03,28),21260,"S-734311","EMISION","VENEGAS CAMPOS VERONICA","110790786",DateTime.Now,"50689795454","C151370"),
        new DailyLog("293912","D006177","VIDA",new DateTime(2025,02,16),new DateTime(2025,03,16),25430,"V-508","RENOVACION","BELLIDO VARGAS JOSE LUIS","502820347",DateTime.Now),
        new DailyLog("343741","VG4220020040731","VIDA GLOBAL ",new DateTime(2025,02,16),new DateTime(2025,03,16),13727,"V-508","EMISION","BELLIDO VARGAS JOSE LUIS","502820347",DateTime.Now),
        new DailyLog("487177","0104ACI0275987-00","ESTUDIANTIL",new DateTime(2025,02,28),new DateTime(2026,02,28),20430,"E-EFECTIVO","EMISION","SOLANO GARRO YANCY MAGALLY","110670140",DateTime.Now,"50685751144")
            };
        }

        private List<ProcedureWithoutMoney> GetProcedures()
        {
            return new List<ProcedureWithoutMoney>()
            {
               new ProcedureWithoutMoney("0204VIA0087142-00","RODRIGUEZ RODRIGUEZ JORGE EDUARDO","CARTA SOLICITUD RESERVA DE DINERO"), 
                new ProcedureWithoutMoney("prueba2-00","RODRIGUEZ RODRIGUEZ JORGE EDUARDO","CARTA SOLICITUD RESERVA DE DINERO")
            };
        }
    }
}
