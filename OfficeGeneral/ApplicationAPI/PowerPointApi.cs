using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficeGeneral.ApplicationAPI
{
    public class PowerPointApi : IApplicationApi
    {
        private readonly PowerPoint.Application _officeApplication;

        public PowerPointApi(PowerPoint.Application application)
        {
            _officeApplication = application;
        }

        public string CurrentDocumentName()
        {
            return _officeApplication.ActivePresentation.Name;
        }
    }
}
