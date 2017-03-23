using Outlook = Microsoft.Office.Interop.Outlook;

namespace OfficeGeneral.ApplicationAPI
{
    public class OutlookApi : IApplicationApi
    {
        private readonly Outlook.Application _officeApplication;

        public OutlookApi(Outlook.Application application)
        {
            _officeApplication = application;
        }

        public string CurrentDocumentName()
        {
            return (_officeApplication.ActiveInspector()?.CurrentItem as Outlook.MailItem)?.Subject;
        }
    }
}
