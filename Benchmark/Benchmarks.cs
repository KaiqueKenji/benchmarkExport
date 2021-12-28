using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using ExportsJuntos.Fakes;
using ExportsJuntos.Models;

namespace Benchmark;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class Benchmarks
{
    private IEnumerable<Portfolio> _portfolios = PortfolioGenerator.GetRandom(10000);
    private static readonly Repository _repository = new Repository();

    [Benchmark(Baseline = true)]
    public void EPPlus()
    {
        _repository.CriarPlanilhaEPPlus(_portfolios, @"C:\dev\Excel\Teste.xlsx");
    }

    [Benchmark]
    public void Mapper()
    {
        _repository.CriarPlanilhaMapper(_portfolios, @"C:\dev\Excel\TesteMapper.xls");
    }

    [Benchmark]
    public void Closed()
    {
        _repository.Write(_portfolios);
    }

    [Benchmark]
    public void Npoi()
    {
        _repository.CriarPlanilhaNpoi(_portfolios);
    }
}