﻿using System;
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

        /// <summary>
        /// .NET Reflector が plug-in をロードした際に呼び出します。
        /// </summary>
        /// <param name="serviceProvider"></param>
        public void Load(IServiceProvider serviceProvider)
        {
            _commandBarManager = serviceProvider.GetTypedService<ICommandBarManager>();
            _assemblyBrowser = serviceProvider.GetTypedService<IAssemblyBrowser>();

            _separator1 = _commandBarManager.CommandBars["Browser.TypeDeclaration"].Items.AddSeparator();
            _button1 = _commandBarManager.CommandBars["Browser.TypeDeclaration"].Items.AddButton(@"Copy Full Assembly Name", CopyFullAssemblyNameButtonClick);

            _separator2 = _commandBarManager.CommandBars["Browser.MethodDeclaration"].Items.AddSeparator();
            _button2 = _commandBarManager.CommandBars["Browser.MethodDeclaration"].Items.AddButton(@"Copy Full Assembly Name", CopyFullAssemblyNameButtonClick);
        }

        private void CopyFullAssemblyNameButtonClick(object sender, EventArgs eventArgs)
        {
            string fullAssemblyName = String.Empty;

            object item = _assemblyBrowser.ActiveItem;
            if (item is ITypeDeclaration)
            {
                ITypeDeclaration typeDeclaration = (ITypeDeclaration) item;
                fullAssemblyName = typeDeclaration.Namespace + "." + typeDeclaration.Name;
            }
            else if (item is IMethodDeclaration)
            {
                IMethodDeclaration methodDeclaration = (IMethodDeclaration) item;
                
                ITypeDeclaration typeDeclaration = methodDeclaration.DeclaringType as ITypeDeclaration;
                if (typeDeclaration != null)
                {
                    fullAssemblyName = typeDeclaration.Namespace + "." + typeDeclaration.Name;
                }
                fullAssemblyName = fullAssemblyName + "." + methodDeclaration.Name;
            }

            if (!String.IsNullOrEmpty(fullAssemblyName))
            {
                Clipboard.SetText(fullAssemblyName);                
            }
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
        }
    }
}