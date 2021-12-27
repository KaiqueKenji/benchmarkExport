using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using ExportsJuntos.Fakes;
using ExportsJuntos.Models;

namespace Benchmark;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class Benchmarks
{
    private IEnumerable<Portfolio> _portfolios = PortfolioGenerator.GetRandom(1);
    private static readonly Repository Teste = new Repository();

    [Benchmark(Baseline = true)]
    public void EPPlus()
    {
        Teste.CriarPlanilhaEPPlus(_portfolios, @"C:\dev\Excel\Teste.xlsx");
    }

    [Benchmark]
    public void Mapper()
    {
        Teste.CriarPlanilhaMapper(_portfolios, @"C:\dev\Excel\TesteMapper.xls");
    }

    [Benchmark]
    public void Closed()
    {
        Teste.Write(_portfolios);
    }

    [Benchmark]
    public void Npoi()
    {
        Teste.CriarPlanilhaNpoi(_portfolios);
    }
}

