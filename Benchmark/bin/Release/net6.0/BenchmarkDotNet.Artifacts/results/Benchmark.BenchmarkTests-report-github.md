``` ini

BenchmarkDotNet=v0.13.1, OS=Windows 10.0.19042.1415 (20H2/October2020Update)
AMD Ryzen 5 2400GE with Radeon Vega Graphics, 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.100
  [Host]     : .NET 6.0.0 (6.0.21.52210), X64 RyuJIT  [AttachedDebugger]
  DefaultJob : .NET 6.0.0 (6.0.21.52210), X64 RyuJIT


```
| Method |      Mean |     Error |    StdDev |    Median | Ratio | RatioSD |     Gen 0 |     Gen 1 |     Gen 2 | Allocated |
|------- |----------:|----------:|----------:|----------:|------:|--------:|----------:|----------:|----------:|----------:|
|   Npoi |  3.447 ms | 0.0662 ms | 0.1800 ms |  3.450 ms |  0.27 |    0.04 |  375.0000 |  125.0000 |         - |      1 MB |
| Mapper |  3.737 ms | 0.0920 ms | 0.2641 ms |  3.650 ms |  0.29 |    0.04 |  257.8125 |   85.9375 |         - |      1 MB |
| EPPlus | 12.891 ms | 0.5855 ms | 1.6514 ms | 12.592 ms |  1.00 |    0.00 | 2000.0000 | 1000.0000 | 1000.0000 |      3 MB |
| Closed | 14.729 ms | 0.3272 ms | 0.9388 ms | 14.217 ms |  1.16 |    0.16 | 1500.0000 |  312.5000 |         - |      4 MB |
