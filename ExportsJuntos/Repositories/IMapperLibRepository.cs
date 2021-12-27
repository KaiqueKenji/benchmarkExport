using ExportsJuntos.Models;

namespace ExportsJuntos.Repositories;

public interface IMapperLibRepository
{
    void CriarPlanilha(IEnumerable<Portfolio> portfolios, string filePath);
}