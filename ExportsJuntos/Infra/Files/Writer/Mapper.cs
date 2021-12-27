using ExportsJuntos.Models;
using ExportsJuntos.Repositories;
using Ganss.Excel;

namespace ExportsJuntos.Infra.Files.Writer;

public class Mapper  : IMapperLibRepository
{

    public void CriarPlanilha(IEnumerable<Portfolio> portfolios, string filePath)
    {
        ExcelMapper excel = new ExcelMapper();

        excel.Save(filePath, portfolios, "TESTE", true);
    }
}