using Word = Microsoft.Office.Interop.Word;

namespace OfficeGeneral.ApplicationAPI
{
    public class WordApi : IApplicationApi
    {
        private readonly Word.Application _officeApplication;

        public WordApi(Word.Application application)
        {
            _officeApplication = application;
        }

        public string CurrentDocumentName()
        {
            return _officeApplication.ActiveDocument.Name;
        }
    }
}
