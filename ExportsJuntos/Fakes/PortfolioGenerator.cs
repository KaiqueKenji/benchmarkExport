using Bogus;
using ExportsJuntos.Models;

namespace ExportsJuntos.Fakes;

public class PortfolioGenerator
{
    public static Portfolio GetRandom() => GetRandom(1).FirstOrDefault();

    public static List<Portfolio> GetRandom(int quantity) =>
        new Faker<Portfolio>()
            .CustomInstantiator(faker => new Portfolio(
                faker.Name.FullName(),
                faker.Name.FullName(),
                faker.Random.Int(),
                faker.Name.FullName(),
                faker.Date.Future(),
                faker.Random.Int(),
                faker.Random.Guid(),
                faker.Random.Guid(),
                faker.Random.Int(),
                faker.Name.FullName(),
                faker.Name.FullName(),
                faker.Date.Future(),
                faker.Random.Decimal(),
                faker.Name.FullName()))
            .Generate(quantity);
}

