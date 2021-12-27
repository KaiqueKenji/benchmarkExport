using BenchmarkDotNet.Attributes;
using ExportsJuntos.Fakes;
using ExportsJuntos.Models;
using ExportsJuntos.Repositories;

namespace ExportsJuntos.Benchmark;

[MemoryDiagnoser]
public class BenchmarkTests
{
    private readonly IMapperLibRepository _mapperLibRepository;
    private readonly IInteropLibRepository _interopLibRepository;
    private readonly IEPPlusLibRepository _epplusLibRepository;
    private readonly IPortfolioFileRepository _portfolioFileRepository;

    private IEnumerable<Portfolio> _portfolios;

    [GlobalSetup]
    public void Setup()
    {
        _portfolios = PortfolioGenerator.GetRandom(1000);
    }

    [Benchmark]
    public void Mapper()
    {
        _mapperLibRepository.CriarPlanilha(_portfolios, @"C:\dev\Excel\TesteMapper.xls");
    }

    //[Benchmark]
    //public void Interop()
    //{
    //    _interopLibRepository.CriarPlanilha(_portfolios, @"C:\dev\Excel\TesteInterop.xls");
    //}

    //[Benchmark]
    //public void EPPlus()
    //{
    //    _epplusLibRepository.CriarPlanilha(_portfolios, @"C:\dev\Excel\TesteEpplus.xlsx");
    //}

    //[Benchmark(Baseline = true)]
    //public void Closed()
    //{
    //    _portfolioFileRepository.Write(_portfolios);
    //}
}