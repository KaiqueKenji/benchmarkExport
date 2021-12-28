``` ini

BenchmarkDotNet=v0.13.1, OS=Windows 10.0.19042.1415 (20H2/October2020Update)
AMD Ryzen 5 2400GE with Radeon Vega Graphics, 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.100
  [Host]     : .NET 6.0.0 (6.0.21.52210), X64 RyuJIT  [AttachedDebugger]
  DefaultJob : .NET 6.0.0 (6.0.21.52210), X64 RyuJIT


```
| Method |     Mean |    Error |   StdDev |   Median | Ratio | RatioSD |        Gen 0 |       Gen 1 |     Gen 2 | Allocated |
|------- |---------:|---------:|---------:|---------:|------:|--------:|-------------:|------------:|----------:|----------:|
| EPPlus |  1.548 s | 0.0589 s | 0.1709 s |  1.470 s |  1.00 |    0.00 |   72000.0000 |  16000.0000 | 5000.0000 |    266 MB |
|   Npoi |  1.681 s | 0.0369 s | 0.1060 s |  1.638 s |  1.10 |    0.13 |  100000.0000 |  24000.0000 | 4000.0000 |    441 MB |
| Mapper |  1.876 s | 0.0591 s | 0.1714 s |  1.813 s |  1.23 |    0.17 |   84000.0000 |  23000.0000 | 3000.0000 |    398 MB |
| Closed | 14.576 s | 0.4891 s | 1.4344 s | 14.520 s |  9.50 |    1.28 | 1042000.0000 | 330000.0000 | 6000.0000 |  3,092 MB |
