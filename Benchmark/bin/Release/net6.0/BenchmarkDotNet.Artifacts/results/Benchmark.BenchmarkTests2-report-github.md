``` ini

BenchmarkDotNet=v0.13.1, OS=Windows 10.0.19042.1415 (20H2/October2020Update)
AMD Ryzen 5 2400GE with Radeon Vega Graphics, 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.100
  [Host]     : .NET 6.0.0 (6.0.21.52210), X64 RyuJIT  [AttachedDebugger]
  DefaultJob : .NET 6.0.0 (6.0.21.52210), X64 RyuJIT


```
| Method |      Mean |     Error |    StdDev |    Median | Ratio | RatioSD |     Gen 0 |     Gen 1 |     Gen 2 | Allocated |
|------- |----------:|----------:|----------:|----------:|------:|--------:|----------:|----------:|----------:|----------:|
| Mapper |  2.794 ms | 0.0553 ms | 0.1560 ms |  2.754 ms |  0.26 |    0.04 |  179.6875 |   74.2188 |         - |    837 KB |
|   Npoi |  3.244 ms | 0.1389 ms | 0.3756 ms |  3.167 ms |  0.30 |    0.06 |         - |         - |         - |  1,286 KB |
| Closed |  5.505 ms | 0.1098 ms | 0.2776 ms |  5.391 ms |  0.50 |    0.08 |  421.8750 |   46.8750 |         - |  1,087 KB |
| EPPlus | 11.122 ms | 0.6341 ms | 1.8195 ms | 10.846 ms |  1.00 |    0.00 | 2000.0000 | 1000.0000 | 1000.0000 |  3,291 KB |
