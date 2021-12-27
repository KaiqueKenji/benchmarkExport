using ExportsJuntos.Infra.Files.Writer;
using ExportsJuntos.Repositories;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddTransient<IPortfolioFileRepository, PortfolioFileRepository>();
builder.Services.AddTransient<IEPPlusLibRepository, EPPlus>();
builder.Services.AddTransient<IInteropLibRepository, Interop>();
builder.Services.AddTransient<IMapperLibRepository, Mapper>();


var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
