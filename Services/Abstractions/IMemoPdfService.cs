namespace MemoGenerator.Services.Abstractions
{
    using System.Threading.Tasks;
    using MemoGenerator.Models;

    public interface IMemoPdfService
    {
        Task<byte[]> GenerateAsync(MemoDocument m);
    }
}
