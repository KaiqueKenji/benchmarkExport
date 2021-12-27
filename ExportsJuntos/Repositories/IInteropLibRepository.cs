using ExportsJuntos.Models;

namespace ExportsJuntos.Repositories;

public interface IInteropLibRepository
{
    void CriarPlanilha(IEnumerable<Portfolio> portfolios, string filePath);
}
