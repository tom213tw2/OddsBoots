using FileHelpers;

namespace OddsBoots.Modal;

[DelimitedRecord(",")]
[IgnoreFirst(1)] 
public class Account
{
    public int WebId { get; set; }
    public string Currency { get; set; }

    public string CompanyWinLoss { get; set; }


    public string Percent { get; set; }

    public string ProfitSharing { get; set; }

    public string Licenes_currency { get; set; }

    public decimal License_EUR { get; set; }
}