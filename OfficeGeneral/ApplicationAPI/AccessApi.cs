using Access = Microsoft.Office.Interop.Access;

namespace OfficeGeneral.ApplicationAPI
{
    public class AccessApi : IApplicationApi
    {
        private readonly Access.Application _officeApplication;

        public AccessApi(Access.Application application)
        {
            _officeApplication = application;
        }

        public string CurrentDocumentName()
        {
            return _officeApplication.CurrentObjectName;
        }
    }
}
