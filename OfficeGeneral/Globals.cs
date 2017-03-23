using OfficeGeneral.ApplicationAPI;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Outlook = Microsoft.Office.Interop.Outlook;
using Access = Microsoft.Office.Interop.Access;

namespace OfficeGeneral
{
    public class Globals
    {
        #region Singleton

        private static readonly object Lock = new object();
        private static volatile Globals _instance;
        public static Globals Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (Lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new Globals();
                        }
                    }
                }
                return _instance;
            }
        }
        private Globals() { }

        #endregion

        #region Office Application

        public static object ApplicationObject { get; set; }

        public IApplicationApi ApplicationApi { get; private set; }

        public IApplicationApi CreateApplicationApi(object application)
        {
            if (ApplicationApi == null)
            {
                ApplicationObject = application;

                if (application is Word.Application)
                {
                    ApplicationApi = new WordApi((Word.Application)application);
                }
                else if (application is Excel.Application)
                {
                    ApplicationApi = new ExcelApi((Excel.Application)application);
                }
                else if (application is PowerPoint.Application)
                {
                    ApplicationApi = new PowerPointApi((PowerPoint.Application)application);
                }
                else if (application is Outlook.Application)
                {
                    ApplicationApi = new OutlookApi((Outlook.Application)application);
                }
                else if (application is Access.Application)
                {
                    ApplicationApi = new AccessApi((Access.Application)application);
                }
            }

            return ApplicationApi;
        }

        #endregion
    }
}
