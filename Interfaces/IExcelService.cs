namespace WebApplication1.Interfaces
{
    public interface IExcelService
    {
        Task<byte[]> ReturnExcelFile();
    }
}
