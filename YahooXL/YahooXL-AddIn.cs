using ExcelDna.Registration;

namespace YahooXL;

public class RtdClockAddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        // Since we have specified ExplicitRegistration=true in the .dna file, we need to do all registration explicitly.
        // Here we only add the async processing, which applies to our IObservable function.
        ExcelRegistration.GetExcelFunctions()
                         .ProcessAsyncRegistrations()
                         .RegisterFunctions();

        ExcelRegistration.GetExcelCommands().RegisterCommands();

        //string xllPath = (string)XlCall.Excel(XlCall.xlGetName);
        //var xlApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
        //xlApp.AddIns.Add(xllPath, false /**don't copy file**/).Installed = true;
        //Debug.WriteLine("autoopen");
    }

    public void AutoClose()
    {
        //Debug.WriteLine("autoclose");
    }

}
