using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeGeneral.ApplicationAPI
{
    public class ExcelApi : IApplicationApi
    {
        private readonly Excel.Application _officeApplication;

        public ExcelApi(Excel.Application application)
        {
            _officeApplication = application;
        }

        public string CurrentDocumentName()
        {
            return _officeApplication.ActiveWorkbook.Name;
        }
    }
}
