using WebApplication1.Models;

namespace WebApplication1.Interfaces
{
    public interface IExcelService
    {
        Task<byte[]> ReturnExcelFile(IEnumerable<DailyLog> dailyLogs, IEnumerable<ProcedureWithoutMoney> procedures, string dailyLogNumber);
    }
}
