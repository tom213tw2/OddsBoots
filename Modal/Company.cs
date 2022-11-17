using FileHelpers;

namespace OddsBoots.Modal;

[DelimitedRecord(",")]
[IgnoreFirst(1)] 
public class Company
{
    public int WebId { get; set; }

    public string BrandName { get; set; }

    public int CompanyAccountId { get; set; }

    public string SocialMediaType { get; set; }

    public string SocialMediaId { get; set; }

    public string IsApi { get; set; }

    public string IsAppEnabled { get; set; }

    public string IsClosed { get; set; }

    public string IsInternal { get; set; }
}