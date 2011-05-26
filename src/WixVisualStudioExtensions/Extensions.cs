using System.Collections.Generic;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;

namespace WixVisualStudioExtensions
{
    public static class Extensions
    {
        public static void AddContextMenuItem(this DTE2 application, AddIn addin, IEnumerable<string> commandBars, string name,
                                               string description, int icon, bool beginGroup, int position)
        {
            var contextGuids = new object[] { };

            var command = application.Commands.AddNamedCommand(addin, name, description, description, false, icon, ref contextGuids,
                            (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled);

            foreach (var commandBarControl in from commandBar in commandBars
                                              select ((CommandBars)application.CommandBars)[commandBar]
                                                  into itemCmdBar
                                                  where itemCmdBar != null
                                                  select (CommandBarControl)command.AddControl(itemCmdBar, position))
            {
                commandBarControl.BeginGroup = beginGroup;
            }
        }

        public static void InsertNewText(this TextSelection selection, string text)
        {
            selection.Insert(text, (int) vsInsertFlags.vsInsertFlagsContainNewText);
        }
    }
}
