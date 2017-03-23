using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using OfficeGeneral.ApplicationAPI;

namespace OfficeAddin.TaskPane
{
    [ComVisible(true)]
    [Guid("66C8F0C3-0CED-42E1-8ED3-6FD8B974DDC2"), ProgId("OfficeAddin.TaskPane.TaskPaneControl")]
    public partial class TaskPaneControl : UserControl
    {
        private CustomTaskPane _taskPane;

        private TaskPaneView _taskPaneView;
        private TaskPaneViewModel _taskPaneViewModel;

        public new bool Visible
        {
            get
            {
                return _taskPane?.Visible ?? false;
            }
            set
            {
                if (_taskPane != null)
                    _taskPane.Visible = value;
            }
        }

        public TaskPaneControl()
        {
            InitializeComponent();
        }

        public void Initialize(CustomTaskPane ctp, IApplicationApi api)
        {
            _taskPane = ctp;

            _taskPaneView = new TaskPaneView();
            _taskPaneViewModel = new TaskPaneViewModel(api) { CloseAction = delegate { Visible = false; } };
            _taskPaneView.DataContext = _taskPaneViewModel;

            elementHost.Child = _taskPaneView;
        }
    }
}
