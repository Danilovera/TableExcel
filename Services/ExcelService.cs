using WebApplication1.Interfaces;
using WebApplication1.Models;

namespace WebApplication1.Services
{
    public class ExcelService : IExcelService
    {
        public Task<byte[]> ReturnExcelFile()
        {
            //TODO: Crear codigo del excel

            foreach (var item in GetDailyLogs())
            {

            }

            foreach (var item in GetProcedures())
            {

            }
            throw new NotImplementedException();
        }


        private List<DailyLog> GetDailyLogs()
        {
            return new List<DailyLog>()
            {
                //TODO: Crear los que faltan
                new DailyLog("343712","0104AUT0258111-01","AUTOS",new DateTime(2025,02,19),new DateTime(2025,03,19),63140,"V-508","RENOVACION","BELLIDO SOLANO LUIS JOSE","402440256",DateTime.Now,"50661393498","CL289812")
            };
        }

        private List<ProcedureWithoutMoney> GetProcedures()
        {
           return new List<ProcedureWithoutMoney>()
            {
               new ProcedureWithoutMoney("0204VIA0087142-00","RODRIGUEZ RODRIGUEZ JORGE EDUARDO","CARTA SOLICITUD RESERVA DE DINERO")
            };
        }
    }
}
