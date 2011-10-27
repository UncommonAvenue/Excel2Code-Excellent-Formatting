using System;
using System.Windows;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.TextManager.Interop;

namespace PasteWithFormatting
{
    class CommandFilter : IOleCommandTarget
    {
        IVsEditorAdaptersFactoryService _adaptersFactory;
        private readonly _DTE _dte;
        IVsTextView _viewAdapter;
        public CommandFilter(_DTE dte, IVsTextView viewAdapter, IVsEditorAdaptersFactoryService adaptersFactory)
        {
            _dte = dte;
            _viewAdapter = viewAdapter;
            _adaptersFactory = adaptersFactory;
        }

        /// <summary>
        /// The next command target in the filter chain (provided by <see cref="IVsTextView.AddCommandFilter"/>).
        /// </summary>
        internal IOleCommandTarget Next { get; set; }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if(pguidCmdGroup == new Guid("{5efc7975-14bc-11cf-9b2b-00aa00573819}") )
            {
                var command = (VSConstants.VSStd97CmdID)nCmdID;
                if (command == VSConstants.VSStd97CmdID.Paste)
                {
                    var data = Clipboard.GetData(DataFormats.CommaSeparatedValue);

                    if (data != null)
                    {

                        Document doc = _dte.ActiveDocument;
                        if (doc != null)
                        {
                            var window = new Window1(data);
                            window.ShowDialog();

                            //If there is no format command, just go skip it all.
                            if(String.IsNullOrEmpty(window.Formatting.Text))
                                return Next.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                            
                            var text = (EnvDTE.TextSelection) doc.Selection;

                            var newText = "";

                            foreach (var line in data.ToString().Replace("\r\n", "\n").Trim('\n').Split('\n'))
                            {
                                try { newText += String.Format(window.Formatting.Text, line.Split(',')) + "\r\n"; }
                                catch (Exception) { }
                            }

                            text.Insert(newText);
                            return VSConstants.S_OK;
                        }
                    }
                }
            }

            return Next.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {

            return Next.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }
    }
}