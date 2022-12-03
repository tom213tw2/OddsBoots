using FileHelpers;

namespace OddsBoots.Modal;

[DelimitedRecord(",")]
[IgnoreFirst(1)] 
public class Account
{
    public int WebId { get; set; }
    public string Currency { get; set; }

    public decimal CompanyWinLoss { get; set; }


    public string Percent { get; set; }
  
    public decimal ProfitSharing { get; set; }

    public decimal Licenes_currency { get; set; }

    public decimal License_EUR { get; set; }
}