using System.Globalization;

namespace OddsBoots.Helper;

public static class DecimalHelper
{
    public static string ConvertStringThousandth(this decimal decimalNumber)
    {
        var str = decimalNumber.ToString(CultureInfo.InvariantCulture).Contains(".00");
        return str ? $"{decimalNumber:N0}" : $"{decimalNumber:N2}";
    }
    
        
}