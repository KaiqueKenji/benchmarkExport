using ExportsJuntos.Models;

namespace ExportsJuntos.Repositories;

public interface IPortfolioFileRepository
{
    Stream Write(IEnumerable<Portfolio> portfolios);
}

