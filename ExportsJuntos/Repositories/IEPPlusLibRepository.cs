using ExportsJuntos.Models;

namespace ExportsJuntos.Repositories;

public interface IEPPlusLibRepository
{
    void CriarPlanilha(IEnumerable<Portfolio> portfolios, string filePath);
}
