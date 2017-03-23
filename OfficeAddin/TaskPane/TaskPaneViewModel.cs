using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using OfficeAddin.WpfApi;
using OfficeGeneral.ApplicationAPI;

namespace OfficeAddin.TaskPane
{
    class TaskPaneViewModel : INotifyPropertyChanged
    {
        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region Properties

        private readonly IApplicationApi _applicationApi;

        #endregion

        #region Bindings

        private string _documentName;
        public string DocumentName
        {
            get { return _documentName; }
            set
            {
                if (value != _documentName)
                {
                    _documentName = value;
                    OnPropertyChanged();
                }
            }
        }

        #endregion

        public TaskPaneViewModel(IApplicationApi api)
        {
            _applicationApi = api;
        }

        #region Commands

        public Action<bool?> CloseAction { get; set; }

        private ICommand _nameCommand;
        public ICommand NameCommand => _nameCommand ?? (_nameCommand = new RelayCommand(GetDocumentName));

        private ICommand _closeCommand;
        public ICommand CloseCommand => _closeCommand ?? (_closeCommand = new RelayCommand(CloseTaskPane));

        #endregion

        private void GetDocumentName()
        {
            DocumentName = _applicationApi.CurrentDocumentName();
        }

        private void CloseTaskPane()
        {
            CloseAction?.Invoke(null);
        }
    }
}
