using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.IO;
using System.Linq;

namespace AngularVSExtension
{
    public static class ExtensionHelper
    {
        public static string GetActiveProjectDirectoryPath(IAsyncServiceProvider ServiceProvider)
        {
            var fullPath = string.Empty;
            try
            {
                var dte2 = (DTE2)ServiceProvider.GetServiceAsync(typeof(Microsoft.VisualStudio.Shell.Interop.SDTE)).Result;
                Array activeSolutionProjects = dte2.ActiveSolutionProjects as Array;
                if (activeSolutionProjects?.Length > 0)
                {
                    var activeProject = activeSolutionProjects.GetValue(0) as Project;
                    fullPath = activeProject.FullName;

                    FileAttributes attr = File.GetAttributes(fullPath);
                    if (!attr.HasFlag(FileAttributes.Directory))
                        fullPath = Path.GetDirectoryName(fullPath);
                }
            }
            finally
            {
                // [ToDo] :: Messagebox to report and Log the Error
            }
            return fullPath;
        }

        public static string GetSelectedFile(IAsyncServiceProvider ServiceProvider)
        {
            var selectedFile = string.Empty;
            try
            {
                var dte2 = (DTE2)ServiceProvider.GetServiceAsync(typeof(Microsoft.VisualStudio.Shell.Interop.SDTE)).Result;

                if (dte2 != null)
                {
                    var selectedItems = (Array)dte2.ToolWindows.SolutionExplorer.SelectedItems;
                    var selectedCollection = selectedItems.Cast<UIHierarchyItem>();

                    foreach (var item in selectedCollection)
                    {
                        selectedFile = item.Object is ProjectItem prjItem
                            ? prjItem.Properties.Item("FullPath").Value.ToString()
                            : string.Empty;
                        break;
                    }
                }
            }
            finally
            {
                // [ToDo] :: Messagebox to report and Log the Error
            }

            return selectedFile;
        }
    }
}