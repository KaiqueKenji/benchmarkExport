``` ini

BenchmarkDotNet=v0.13.1, OS=Windows 10.0.19042.1415 (20H2/October2020Update)
AMD Ryzen 5 2400GE with Radeon Vega Graphics, 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.100
  [Host]     : .NET 6.0.0 (6.0.21.52210), X64 RyuJIT  [AttachedDebugger]
  DefaultJob : .NET 6.0.0 (6.0.21.52210), X64 RyuJIT


```
| Method |      Mean |     Error |    StdDev |    Median | Ratio | RatioSD |     Gen 0 |     Gen 1 |     Gen 2 | Allocated |
|------- |----------:|----------:|----------:|----------:|------:|--------:|----------:|----------:|----------:|----------:|
|   Npoi |  2.972 ms | 0.0594 ms | 0.1446 ms |  2.950 ms |  0.26 |    0.04 |         - |         - |         - |  1,286 KB |
| Mapper |  3.185 ms | 0.1549 ms | 0.4444 ms |  3.103 ms |  0.29 |    0.05 |  183.5938 |   70.3125 |         - |    837 KB |
| Closed |  7.852 ms | 0.3985 ms | 1.0775 ms |  7.512 ms |  0.70 |    0.12 |         - |         - |         - |  1,132 KB |
| EPPlus | 11.297 ms | 0.5650 ms | 1.5561 ms | 11.020 ms |  1.00 |    0.00 | 2000.0000 | 1000.0000 | 1000.0000 |  3,291 KB |
