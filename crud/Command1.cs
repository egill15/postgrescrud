using System;
using System.CodeDom.Compiler;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace crud
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command1
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("ab7ab883-cece-4e07-bd80-00d383e17bc3");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Command1(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command1 Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
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
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command1's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new Command1(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        /// 


        private static EnvDTE80.DTE2 GetDTE2()
        {
            return Package.GetGlobalService(typeof(DTE)) as EnvDTE80.DTE2;
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            DTE dte = Package.GetGlobalService(typeof(SDTE)) as DTE;
            if (dte.SelectedItems.Count <= 0) return;

            foreach (SelectedItem selectedItem in dte.SelectedItems)
            {
                if (selectedItem.ProjectItem == null) return;
                var projectItem = selectedItem.ProjectItem;
                var fullPathProperty = projectItem.Properties.Item("FullPath");
                if (fullPathProperty == null) return;
                var fullPath = fullPathProperty.Value.ToString();
                GetProperties(fullPath);
            }
        }

        private void GetProperties(string location)
        {
            FileInfo file = new FileInfo(location);

            if (file.Extension.ToUpper(CultureInfo.InvariantCulture) != ".CS")
            {
                return;
            }

            CompilerParameters cp = new CompilerParameters();
            cp.GenerateExecutable = false;
            cp.GenerateInMemory = true;
            cp.TreatWarningsAsErrors = false;
            cp.ReferencedAssemblies.Add("System.Data.dll");
            cp.ReferencedAssemblies.Add("System.dll");
            cp.ReferencedAssemblies.Add("System.Xml.dll");
            cp.ReferencedAssemblies.Add("System.Linq.dll");
            cp.ReferencedAssemblies.Add("System.Xml.Linq.dll");
            cp.ReferencedAssemblies.Add("System.Core.dll");
            cp.ReferencedAssemblies.Add("System.ComponentModel.dll");
            cp.ReferencedAssemblies.Add("System.ComponentModel.DataAnnotations.dll");

            CompilerResults cr = CodeDomProvider.CreateProvider("CSharp").CompileAssemblyFromFile(cp, location);

            if (cr.Errors.Count > 0)
            {
                MessageBox.Show(string.Format("ERROR, must select a model file"));
                foreach (CompilerError ce in cr.Errors)
                {
                    MessageBox.Show(string.Format(ce.ToString()));
                }
            }
            else
            {
                var classItems = cr.CompiledAssembly.GetTypes().ToList();

                foreach (var classItem in classItems)
                {
                    GenerateCrud(classItem);
                }
            }
        }

        public void GenerateCrud(Type classname)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(Add(classname));
            sb.Append("\n");
            sb.Append(FindAll(classname));
            sb.Append("\n");
            sb.Append(FindById(classname));
            sb.Append("\n");
            sb.Append(Remove(classname));
            sb.Append("\n");
            sb.Append(Update(classname));

            Clipboard.SetText(sb.ToString());

        }

        public string GetPropsWithoutId(Type classname)
        {
            var props = classname.GetProperties();
            StringBuilder sb = new StringBuilder();
            foreach (var prop in props)
            {
                if (prop.Name != "Id")
                {
                    sb.Append(prop.Name.ToLower());
                    if (props.Last().Name != prop.Name)
                    {
                        sb.Append(", ");
                    }
                }
            }
            return sb.ToString();
        }

        public string GetPropsWithAttWithoutId(Type classname)
        {
            var props = classname.GetProperties();
            StringBuilder sb = new StringBuilder();
            foreach (var prop in props)
            {
                if (prop.Name != "Id")
                {
                    sb.Append("@" + prop.Name.ToLower());
                    if (props.Last().Name != prop.Name)
                    {
                        sb.Append(", ");
                    }
                }
            }
            return sb.ToString();
        }

        public string GetProps(Type classname)
        {
            var props = classname.GetProperties();
            StringBuilder sb = new StringBuilder();
            foreach (var prop in props)
            {
                sb.Append(prop.Name.ToLower());
                if (props.Last().Name != prop.Name)
                {
                    sb.Append(", ");
                }
                
            }
            return sb.ToString();
        }

        public string GetPropsWithAtt(Type classname)
        {
            var props = classname.GetProperties();
            StringBuilder sb = new StringBuilder();
            foreach (var prop in props)
            {
                sb.Append("@" + prop.Name.ToLower());
                if (props.Last().Name != prop.Name)
                {
                    sb.Append(", ");
                }
            }
            return sb.ToString();
        }

        public string Add(Type classname)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("public int Add(" + classname.Name + " item)");
            sb.Append("\n{");
            sb.Append("\n\tusing IDbConnection dbConnection = Connection;");
            sb.Append("\n\tdbConnection.Open();");
            sb.Append("\n\treturn dbConnection.QueryFirst<int>(\"INSERT INTO " + classname.Name.ToLower() + " (" + GetPropsWithoutId(classname) + ") VALUES(" + GetPropsWithAttWithoutId(classname) + ") RETURNING id\", item);");
            sb.Append("\n}");
            return sb.ToString();
        }

        public string FindAll(Type classname)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("public IEnumerable<" + classname.Name + "> FindAll()");
            sb.Append("\n{");
            sb.Append("\n\tusing IDbConnection dbConnection = Connection;");
            sb.Append("\n\tdbConnection.Open();");
            sb.Append("\n\treturn dbConnection.Query<" + classname.Name + ">(\"SELECT " + GetProps(classname) + " FROM " +classname.Name.ToLower() + "\");");
            sb.Append("\n}");
            return sb.ToString();
        }

        public string FindById(Type classname)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("public " + classname.Name + " FindById(int id)");
            sb.Append("\n{");
            sb.Append("\n\tusing IDbConnection dbConnection = Connection;");
            sb.Append("\n\tdbConnection.Open();");
            sb.Append("\n\treturn dbConnection.Query<" + classname.Name + ">(\"SELECT " + GetProps(classname) + " FROM " + classname.Name.ToLower() + " WHERE id = @ID\", new { ID = id }).FirstOrDefault();");
            sb.Append("\n}");
            return sb.ToString();
        }

        public string Remove(Type classname)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("public int Remove(int id)");
            sb.Append("\n{");
            sb.Append("\n\tusing IDbConnection dbConnection = Connection;");
            sb.Append("\n\tdbConnection.Open();");
            sb.Append("\n\tdbConnection.Execute(\"DELETE FROM " + classname.Name.ToLower() + " WHERE id = @Id\", new { Id = id });");
            sb.Append("\n\treturn id;");
            sb.Append("\n}");
            return sb.ToString();
        }

        public string Update(Type classname)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("public int Update(" + classname.Name + " item)");
            sb.Append("\n{");
            sb.Append("\n\tusing IDbConnection dbConnection = Connection;");
            sb.Append("\n\tdbConnection.Open();");
            sb.Append("\n\tdbConnection.Execute(\"UPDATE " + classname.Name.ToLower() + " SET " + UpdateQuery(classname) + " WHERE id = @Id\", item);");
            sb.Append("\n\treturn item.Id;");
            sb.Append("\n}");
            return sb.ToString();
        }

        public string UpdateQuery(Type classname)
        {
            var props = classname.GetProperties();
            StringBuilder sb = new StringBuilder();
            foreach (var prop in props)
            {
                if(prop.Name != "Id")
                {
                    sb.Append(prop.Name.ToLower() + " = @" + prop.Name);
                    if (props.Last().Name != prop.Name)
                    {
                        sb.Append(", ");
                    }
                }
            }
            return sb.ToString();
        }
    }
}
