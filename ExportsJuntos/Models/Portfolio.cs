using ExportsJuntos.Entities;

namespace ExportsJuntos.Models;

public class Portfolio : Entity
{
    public Portfolio(
        string portfolioName,
        string description,
        int currencyId,
        string assetName,
        DateTime date,
        int portfolioTypeId,
        Guid tenantId,
        Guid managerId,
        int countryId,
        string name,
        string nickname,
        DateTime birthDate,
        decimal age,
        string document)
    {
        PortfolioName = portfolioName;
        Description = description;
        CurrencyId = currencyId;
        AssetName = assetName;
        Date = date;
        PortfolioTypeId = portfolioTypeId;
        TenantId = tenantId;
        ManagerId = managerId;
        CountryId = countryId;
        Name = name;
        Nickname = nickname;
        BirthDate = birthDate;
        Age = age;
        Document = document;
    }

    public string PortfolioName { get; set; }
    public string Description { get; set; }
    public int CurrencyId { get; set; }
    public string AssetName { get; set; }
    public DateTime Date { get; set; }
    public int PortfolioTypeId { get; set; }
    public Guid TenantId { get; set; }
    public Guid ManagerId { get; set; }
    public int CountryId { get; set; }
    public string Name { get; set; }
    public string Nickname { get; set; }
    public DateTime BirthDate { get; set; }
    public decimal Age { get; set; }
    public string Document { get; set; }
}