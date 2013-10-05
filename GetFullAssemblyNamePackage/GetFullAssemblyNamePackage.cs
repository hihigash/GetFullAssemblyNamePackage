using System;
using System.Diagnostics;
using System.Windows.Forms;
using Reflector.Application;
using Reflector.CodeModel;

namespace Reflector.AddIns
{
    public class GetFullAssemblyNamePackage : IPackage
    {
        private ICommandBarManager _commandBarManager;
        private IAssemblyBrowser _assemblyBrowser;

        private ICommandBarSeparator _separator1;
        private ICommandBarButton _button1;
        private ICommandBarSeparator _separator2;
        private ICommandBarButton _button2;
        private ICommandBarSeparator _separator3;
        private ICommandBarButton _button3;

        /// <summary>
        /// .NET Reflector が plug-in をロードした際に呼び出します。
        /// </summary>
        /// <param name="serviceProvider"></param>
        public void Load(IServiceProvider serviceProvider)
        {
            _commandBarManager = serviceProvider.GetTypedService<ICommandBarManager>();
            _assemblyBrowser = serviceProvider.GetTypedService<IAssemblyBrowser>();

            _assemblyBrowser.ActiveItemChanged += AssemblyBrowserActiveItemChanged;

            _separator1 = _commandBarManager.CommandBars["Browser.TypeDeclaration"].Items.AddSeparator();
            _button1 = _commandBarManager.CommandBars["Browser.TypeDeclaration"].Items.AddButton(@"Copy Full Assembly Name", CopyFullAssemblyNameButtonClick);

            _separator2 = _commandBarManager.CommandBars["Browser.MethodDeclaration"].Items.AddSeparator();
            _button2 = _commandBarManager.CommandBars["Browser.MethodDeclaration"].Items.AddButton(@"Copy Full Assembly Name", CopyFullAssemblyNameButtonClick);

            _separator3 = _commandBarManager.CommandBars["Edit"].Items.AddSeparator();
            _button3 = _commandBarManager.CommandBars["Edit"].Items.AddButton(@"Copy Full Assembly Name", CopyFullAssemblyNameButtonClick, Keys.Alt|Keys.Control|Keys.O);
        }

        /// <summary>
        /// Assembly Browser の要素が変更された場合に呼び出されます。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void AssemblyBrowserActiveItemChanged(object sender, EventArgs e)
        {
            // サポートしていないタイプの要素だった場合はボタンを無効にする。
            object item = _assemblyBrowser.ActiveItem;
            _button3.Enabled = (item is ITypeDeclaration || item is IMethodDeclaration);
        }

        private void CopyFullAssemblyNameButtonClick(object sender, EventArgs eventArgs)
        {
            string fullAssemblyName = String.Empty;

            object item = _assemblyBrowser.ActiveItem;
            if (item is ITypeDeclaration)
            {
                ITypeDeclaration typeDeclaration = (ITypeDeclaration) item;
                fullAssemblyName = String.Format("{0}.{1}", typeDeclaration.Namespace, typeDeclaration.Name);
            }
            else if (item is IMethodDeclaration)
            {
                IMethodDeclaration methodDeclaration = (IMethodDeclaration) item;
                ITypeDeclaration typeDeclaration = (ITypeDeclaration)methodDeclaration.DeclaringType;
                fullAssemblyName = String.Format("{0}.{1}.{2}", typeDeclaration.Namespace, typeDeclaration.Name, methodDeclaration.Name);
            }

            if (!String.IsNullOrEmpty(fullAssemblyName))
            {
                Clipboard.SetText(fullAssemblyName);                
            }

            Debug.WriteLine(fullAssemblyName);
        }

        /// <summary>
        /// .NET Reflector が plug-in をアンロードした際に呼び出します。
        /// </summary>
        public void Unload()
        {
            _commandBarManager.CommandBars["Browser.TypeDeclaration"].Items.Remove(_separator1);
            _commandBarManager.CommandBars["Browser.TypeDeclaration"].Items.Remove(_button1);

            _commandBarManager.CommandBars["Browser.MethodDeclaration"].Items.Remove(_separator2);
            _commandBarManager.CommandBars["Browser.MethodDeclaration"].Items.Remove(_button2);

            _commandBarManager.CommandBars["Edit"].Items.Remove(_separator3);
            _commandBarManager.CommandBars["Edit"].Items.Remove(_button3);
        }
    }
}
