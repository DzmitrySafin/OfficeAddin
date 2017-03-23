using System;
using System.Resources;
using System.Runtime.InteropServices;
using System.Windows;
using Extensibility;
using Microsoft.Office.Core;
using OfficeAddin.Properties;
using OfficeAddin.TaskPane;
using OfficeGeneral;
using OfficeGeneral.ApplicationAPI;

namespace OfficeAddin
{
    [Guid("FC098D9B-7077-47A7-8F95-9BDAC9B6915F"), ProgId("OfficeAddin.Connect")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        #region Properties

        private IRibbonUI _ribbon;
        private readonly ResourceManager _resourceManager = new ResourceManager(typeof(Resources));

        private ICTPFactory _ctpFactory;
        private TaskPaneControl _taskPaneControl;

        private IApplicationApi _applicationApi;

        #endregion

        #region IDTExtensibility2

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationApi = Globals.Instance.CreateApplicationApi(application);
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public void OnStartupComplete(ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            //throw new NotImplementedException();
        }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonId)
        {
            return Resources.Ribbon;
        }

        #endregion

        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(ICTPFactory ctpFactoryInst)
        {
            _ctpFactory = ctpFactoryInst;
        }

        #endregion

        #region Task Pane

        private TaskPaneControl CreateCustomTaskPane()
        {
            CustomTaskPane ctp = _ctpFactory.CreateCTP("OfficeAddin.TaskPane.TaskPaneControl", Resources.TaskPaneHeader);
            ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            ctp.VisibleStateChange += TaskPane_OnVisibleStateChange;

            var taskPane = (TaskPaneControl) ctp.ContentControl;
            taskPane.Initialize(ctp, _applicationApi);
            return taskPane;
        }

        private void TaskPane_OnVisibleStateChange(CustomTaskPane customTaskPaneInst)
        {
            _ribbon.InvalidateControl("btnTaskPane");
        }

        #endregion

        #region Ribbon - General

        public void Ribbon_OnLoad(IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        public string Ribbon_GetLabel(IRibbonControl control)
        {
            return _resourceManager.GetString(control.Id + "_Label");
        }

        #endregion

        #region Ribbon - Buttons

        public void ButtonUsd_OnClick(IRibbonControl control)
        {
            MessageBox.Show(_applicationApi.CurrentDocumentName(), control.Id);
        }

        public void ButtonEur_OnClick(IRibbonControl control)
        {
            MessageBox.Show(_applicationApi.CurrentDocumentName(), control.Id);
        }

        public void ButtonByn_OnClick(IRibbonControl control)
        {
            MessageBox.Show(_applicationApi.CurrentDocumentName(), control.Id);
        }

        public bool ButtonTaskPane_GetPressed(IRibbonControl control)
        {
            return _taskPaneControl != null && _taskPaneControl.Visible;
        }

        public void ButtonTaskPane_OnClick(IRibbonControl control, bool pressed)
        {
            if (_taskPaneControl == null)
            {
                _taskPaneControl = CreateCustomTaskPane();
            }

            if (_taskPaneControl != null)
            {
                _taskPaneControl.Visible = pressed;
            }
        }

        #endregion
    }
}
