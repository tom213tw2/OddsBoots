using FileHelpers;

namespace OddsBoots.Helper;

public class FileHelper
{
    public static bool Get<T>(string aFileName, out T[] aoDataList, out string aErrorMsg)
    {
        aErrorMsg = "";
        aoDataList = Array.Empty<T>();


        if (!File.Exists(aFileName))
        {
            aErrorMsg += $"檔案不存在!!檔案名稱:{aFileName}\n";
            return false;
        }

        var engine = new FileHelperEngine<T>
        {
            ErrorManager =
            {
                ErrorMode = ErrorMode.SaveAndContinue
            }
        };

        aoDataList = engine.ReadFile(aFileName);

        if (engine.ErrorManager.ErrorCount <= 0) return true;
        foreach (var err in engine.ErrorManager.Errors)
        {
            aErrorMsg += $"檔案讀取時發生錯誤!!檔案名稱:{aFileName}\n";
            aErrorMsg += $"錯誤行號: {err.LineNumber}\n";
            aErrorMsg += $"錯誤原因: {err.RecordString}\n";
            aErrorMsg += $"完整錯誤訊息: {err.ExceptionInfo}";
            break;
        }
        return false;

    }
}