using BenchmarkDotNet.Running;
using ExportsJuntos.Benchmark;
using ExportsJuntos.Fakes;
using ExportsJuntos.Infra.Files.Writer;
using ExportsJuntos.Repositories;
using ExportsJuntos.Results;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace ExportsJuntos.Controllers
{

    [ApiController]
    [Route("[controller]")]
    public class PortfolioController : Controller
    {
        private readonly ILogger<PortfolioController> _logger;
        private readonly IPortfolioFileRepository _portfolioFileRepository;
        private readonly IEPPlusLibRepository _epplusLibRepository;
        private readonly IInteropLibRepository _interopLibRepository;
        private readonly IMapperLibRepository _mapperLibRepository;

        public PortfolioController(
            ILogger<PortfolioController> logger,
            IPortfolioFileRepository portfolioFileRepository,
            IEPPlusLibRepository epplusLibRepository,
            IInteropLibRepository interopLibRepository,
            IMapperLibRepository mapperLibRepository)
        {
            _logger = logger;
            _portfolioFileRepository = portfolioFileRepository;
            _epplusLibRepository = epplusLibRepository;
            _interopLibRepository = interopLibRepository;
            _mapperLibRepository = mapperLibRepository;
        }

        /// <summary>
        /// Método que adiciona um portfolio
        /// </summary>
        [ProducesResponseType(StatusCodes.Status201Created)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [HttpGet("v1/benchmark")]
        public void RunBenchmark()
        {
            BenchmarkRunner.Run<BenchmarkTests>();
        }

            /// <summary>
            /// Método que adiciona um portfolio
            /// </summary>
            /// <returns>Id</returns>
            [ProducesResponseType(StatusCodes.Status201Created, Type = typeof(PortfolioResult))]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [HttpGet("v1/portfolios")]
        public IActionResult ClosedXmlLib([FromQuery] string outputFormat)
        {
            var portfolios = PortfolioGenerator.GetRandom(100);

            return outputFormat switch
            {
                "closed" => ReturnClosed(),
                "epplus" => ReturnEpplus(),
                "interop" => ReturnInterop(),
                "mapper" => ReturnMapper(),
                _ => BadRequest(nameof(outputFormat), "Formato não suportado")
            };

            IActionResult ReturnClosed()
            {
                var stream = _portfolioFileRepository.Write(portfolios);
                if (stream is null) return NoContent();
                return File(
                    fileStream: stream,
                    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileDownloadName: $"portfolios.xlsx");
            }

            IActionResult ReturnEpplus()
            {
                var filePath = @"C:\dev\Excel\Teste.xlsx";
                _epplusLibRepository.CriarPlanilha(portfolios, filePath);
                return Ok("Excel criado com Sucesso!");
            }

            IActionResult ReturnInterop()
            {
                var teste = @"C:\dev\Excel\TesteInterop.xls";
                _interopLibRepository.CriarPlanilha(portfolios, teste);
                return Ok("Excel criado com Sucesso!");
            }

            IActionResult ReturnMapper()
            {
                var teste = @"C:\dev\Excel\TesteMapper.xls";
                _interopLibRepository.CriarPlanilha(portfolios, teste);
                return Ok("Excel criado com Sucesso!");
            }
        }

        [HttpGet]
        [Route("export")]
        [ProducesResponseType((int)HttpStatusCode.OK)]
        public async Task<FileResult> Export(string cookieName = null)
        {
            var byteArray = await NpoiLib.CreateExcelFileAsync();

            return this.DownloadExcelFile(byteArray, "MyFile.xlsx", cookieName);
        }
    }
}

public static class ControllerExtensions
{
    public static FileResult DownloadExcelFile(this ControllerBase controller, byte[] byteArray, string fileName, string cookieName)
    {
        if (!string.IsNullOrEmpty(cookieName))
            controller.Response.Cookies.Append(cookieName, "true");

        return controller.File(byteArray, "application/Excel", fileName);
    }
}