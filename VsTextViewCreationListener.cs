using System.ComponentModel.Composition;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;

namespace PasteWithFormatting
{
    [Export(typeof(IVsTextViewCreationListener))]
    [ContentType("any")]
    [TextViewRole(PredefinedTextViewRoles.Interactive)]
    class VsTextViewCreationListener : IVsTextViewCreationListener
    {
        private _DTE _dte;

        [Import]
        IVsEditorAdaptersFactoryService AdaptersFactory = null;

        [ImportingConstructor]
        public VsTextViewCreationListener(SVsServiceProvider serviceProvider)
        {
            _dte = (_DTE)serviceProvider.GetService(typeof(_DTE));
        }
      

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            CommandFilter filter = new CommandFilter(_dte, textViewAdapter, AdaptersFactory);

            IOleCommandTarget next;
            if (ErrorHandler.Succeeded(textViewAdapter.AddCommandFilter(filter, out next)))
                filter.Next = next;
        }
    }
}