//------------------------------------------------------------------------------
// <copyright file="MenuCmd.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using EnvDTE80;
using System.Diagnostics;
using EnvDTE;
using System.Text.RegularExpressions;


namespace ClassRefactorCmd
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class MenuCmd
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("08bf97ab-53e6-4c1a-b01b-69d371718a49");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="MenuCmd"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private MenuCmd(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static MenuCmd Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new MenuCmd(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            
            DTE2 dte = ServiceProvider.GetService(typeof(DTE)) as DTE2;
            if(dte.Solution.FullName == null || dte.Solution.FullName == "")
            {
                Debug.WriteLine("no opened solution");
                return;
            }

            //one can use FindProjectItem to find a file in opened solution.
            string filePath = "test.cs";
            //filePath can also be a partial path like "Folder1\\SubFolder\\test.cs";
            ProjectItem item = dte.Solution.FindProjectItem(filePath);
            int numScriptRenamedSucceed = 0;
            if (item != null)
            {
                //use regular expression to recognize csharp script file
                Match match = Regex.Match(item.FileNames[0], @".+\.cs$");
                if(match.Success)
                {
                    //if there is a class call TestClass in test.cs, and one would like to change the name to NewTestClass
                    if(renameClassName(item.FileCodeModel.CodeElements, "TestClass", "NewTestClass"))
                    {
                        Debug.WriteLine("renaming succeed");
                        numScriptRenamedSucceed++;
                    }
                    else
                    {
                        Debug.WriteLine("renaming failed");
                    }
                }                
            }
            else
            {
                Debug.WriteLine("file not found");
            }

            //if script was successfully saved, then save all projects.(Of course, one can only save the project that have been changed)
            if(numScriptRenamedSucceed > 0)
            {
                //I didn't find a way to save solution directly with old solution name.
                //so I just use the save function provided in project level.
                Projects projects = dte.Solution.Projects;
                foreach(Project project in projects)
                {
                    try
                    {
                        project.Save(); //save with old project name
                    }
                    catch(Exception exp)
                    {

                    }
                }
            }
            
        }

        bool renameClassName(CodeElements elements, string oldClassName, string newClassName)
        {
            bool changingNameFailed = false;
            foreach(CodeElement element in elements)
            {
                try
                {
                    CodeElement2 element2 = element as CodeElement2;
                    if((element2.Kind == vsCMElement.vsCMElementClass //it's a class node in syntax tree
                        //|| element2.Kind == vsCMElement.vsCMElementInterface //it's an interface node
                        ) && element2.Name == oldClassName) 
                    {
                        changingNameFailed = true;
                        element2.RenameSymbol(newClassName);
                        //if renaming succeed, program will reach here without exception being caught.
                        return true;
                    }
                }
                catch(Exception e)
                {
                    if(changingNameFailed)
                    {
                        return false;
                    }

                    //Debug.WriteLine(e.ToString());
                }

                if(element.Kind == vsCMElement.vsCMElementNamespace)
                {
                    if(renameClassName(((CodeNamespace)element).Members, oldClassName, newClassName))
                    {
                        return true;
                    }
                    //else continue
                }
                else if(element.IsCodeType)
                {
                    if(renameClassName(((CodeType)element).Members, oldClassName, newClassName))
                    {
                        return true;
                    }
                    //else continue
                }
            }

            return false; //if it didn't match any class name
        }

        //Traverse every file in a solution.
        void TraverseAllItems(Solution solution)
        {
            Projects projects = solution.Projects; //one solution can contain several projects
            foreach(Project project in projects)
            {
                //one project may have several files(ProjectItems)
                ProjectItems items = project.ProjectItems;
                foreach(ProjectItem item in items)
                {
                    //usually one item would only contain one file path(file name)
                    Debug.WriteLine("file path:" + item.FileNames[0]);
                }
            }
        }
    }
}
