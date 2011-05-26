using System;
using EnvDTE;
using EnvDTE80;
using Extensibility;

namespace WixVisualStudioExtensions
{
    public class Addin : IDTExtensibility2, IDTCommandTarget
    {
        private const string LocationCodeWindow = "Code Window";
        private const string CommandRoot = "WixVisualStudioExtensions.Addin.";
        private const string InsertGuidCommand = "InsertGuid";
        private const string InsertFileIdCommand = "InsertFileId";
        private const string InsertShortFilenameCommand = "InsertShortFileName";
        private const string InsertDirectoryElementCommand = "InsertDirectoryElement";
        private const string InsertComponentElementCommand = "InsertComponentElement";
        private const string InsertFileElementCommand = "InsertFileElement";
        private const string InsertFileAndShortcutElementCommand = "InsertFileAndShortcutElement";

        private DTE2 _applicationObject;
        private AddIn _addInInstance;

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = (DTE2)application;
            _addInInstance = (AddIn)addInInst;

            if (connectMode != ext_ConnectMode.ext_cm_UISetup) return;

            _applicationObject.AddContextMenuItem(_addInInstance, new[] { LocationCodeWindow }, InsertGuidCommand, "Insert WiX Guid", 101, true, 101);
            _applicationObject.AddContextMenuItem(_addInInstance, new[] { LocationCodeWindow }, InsertFileIdCommand, "Insert WiX Id", 101, false, 102);
            _applicationObject.AddContextMenuItem(_addInInstance, new[] { LocationCodeWindow }, InsertShortFilenameCommand, "Insert WiX Short Filename", 101, false, 103);
            _applicationObject.AddContextMenuItem(_addInInstance, new[] { LocationCodeWindow }, InsertDirectoryElementCommand, "Insert WiX Directory", 101, false, 104);
            _applicationObject.AddContextMenuItem(_addInInstance, new[] { LocationCodeWindow }, InsertComponentElementCommand, "Insert WiX Component", 101, false, 105);
            _applicationObject.AddContextMenuItem(_addInInstance, new[] { LocationCodeWindow }, InsertFileElementCommand, "Insert WiX File", 101, false, 106);
            _applicationObject.AddContextMenuItem(_addInInstance, new[] { LocationCodeWindow }, InsertFileAndShortcutElementCommand, "Insert WiX File And Shortcut", 101, false, 107);
        }

        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            if (neededText != vsCommandStatusTextWanted.vsCommandStatusTextWantedNone || !commandName.StartsWith(CommandRoot)) return;
            status = vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
            return;
        }

        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            var selection = _applicationObject.ActiveDocument.Selection as TextSelection;
            if (executeOption != vsCommandExecOption.vsCommandExecOptionDoDefault || selection == null) return;
            handled = true;
            switch (commandName)
            {
                case CommandRoot + InsertGuidCommand: selection.Text = WixContent.GenerateNewGuid(); break;
                case CommandRoot + InsertFileIdCommand: selection.Text = WixContent.GenerateFileId(); break;
                case CommandRoot + InsertShortFilenameCommand: selection.Text = WixContent.GenerateShortFileName(); break;
                case CommandRoot + InsertDirectoryElementCommand: selection.InsertNewText(WixContent.GenerateDirectoryTag()); break;
                case CommandRoot + InsertComponentElementCommand: selection.InsertNewText(WixContent.GenerateComponentTag()); break;
                case CommandRoot + InsertFileElementCommand: selection.InsertNewText(WixContent.GenerateFileTag(false)); break;
                case CommandRoot + InsertFileAndShortcutElementCommand: selection.InsertNewText(WixContent.GenerateFileTag(true)); break;
                default: handled = false; break;
            }
        }

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
    }
}