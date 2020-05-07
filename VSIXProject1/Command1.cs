using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace VSIXProject1
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
		public static readonly Guid CommandSet = new Guid("5dfaffe0-48c0-47bd-b512-5e42179db5df");

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

			var menuCommandId = new CommandID(CommandSet, CommandId);
			var menuItem = new MenuCommand(this.Execute, menuCommandId);
			commandService.AddCommand(menuItem);
		}

		private static DTE _dte;

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

			_dte = (EnvDTE.DTE)await package.GetServiceAsync(typeof(EnvDTE.DTE));

			Instance = new Command1(package, commandService);

		}

		/// <summary>
		/// This function is the callback used to execute the command when the menu item is clicked.
		/// See the constructor to see how the menu item is associated with this function using
		/// OleMenuCommandService service and MenuCommand class.
		/// </summary>
		/// <param name="sender">Event sender.</param>
		/// <param name="e">Event args.</param>
		private void Execute(object sender, EventArgs e)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
			string title = "Command1";

			
			var projects = GetEnvDteProjects(_dte);

			
			foreach (Project project in projects)
			{
				string projectName = "ClassLibrary1";
				string SolutionFolderName = "Libs";

				if (project.Name == projectName)
				{
					this.SafeMoveProjectToSolutionFolder(projects, project, projectName, SolutionFolderName);
				}
			}
		}

		private void SafeMoveProjectToSolutionFolder(List<Project> projects, Project project, string projectName, string SolutionFolderName)
		{
			var projectsToAddRefBackTo = new System.Collections.Generic.List<Project>();
			foreach (Project project2 in projects)
			{
				if (this.HasReferenceTo(project2, projectName))
				{
					projectsToAddRefBackTo.Add(project2);
				}
			}

			Project movedProject = this.AddProjectToSolutionFolder(project, SolutionFolderName);

			foreach (Project project1 in projectsToAddRefBackTo)
			{
				this.AddProjectReference(project1, movedProject);
			}
		}

		private void AddProjectReference(Project baseProject, Project projectToReference)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			if (baseProject.Object is VSLangProj.VSProject vsProject)
			{
				VSLangProj.Reference reference = null;
				try
				{
					reference = vsProject.References.Find(projectToReference.Name);
				}
				catch (Exception ex)
				{
					//reference doesnt exist, good to go
				}

				if (reference != null)
				{

					throw new InvalidOperationException("Reference already exists.");
				}

				vsProject.References.AddProject(projectToReference);
			}
		}

		private Project AddProjectToSolutionFolder(Project project1, string folderName)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			try
			{
				EnvDTE80.Solution2 solution = (EnvDTE80.Solution2)((EnvDTE80.DTE2)_dte).Solution;

				string projFullname = project1.FullName;

				Project project = GetSolutionSubFolder(solution, folderName);

				if (project != null)
				{
					EnvDTE80.SolutionFolder folder = (EnvDTE80.SolutionFolder)project.Object;
					solution.Remove(project1);
					project1 = folder.AddFromFile(projFullname);
				}
			}
			catch (Exception e)
			{
			}
			return project1;
		}

		private static Project GetSolutionSubFolder(EnvDTE80.Solution2 solution, string folderName)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Projects projects = solution.Projects;
			Project folder = projects.Cast<Project>().FirstOrDefault(p => string.Equals(p.Name, folderName));

			if (folder == null)
			{
				folder = solution.AddSolutionFolder(folderName);
			}

			return folder;
		}

		private System.Collections.Generic.Dictionary<string, string> GetProjectReferences(Project project)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var refs = new System.Collections.Generic.Dictionary<string, string>();
			if (project.Object is VSLangProj.VSProject vsProject)
			{
				foreach (VSLangProj.Reference reference in vsProject.References)
				{
					var identity = reference.Identity;
					string path = "";
					if (reference.StrongName)
					{
						path = $"{reference.Identity} , Version={reference.Version}, Culture={(ParseReferenceCulture(reference))}, PublicKeyToken={reference.PublicKeyToken}";
					}
					else
					{
						path = reference.Path;
					}

					refs.Add(identity, path);
				}
			}

			return refs;

			string ParseReferenceCulture(VSLangProj.Reference reference)
			{
				return string.IsNullOrEmpty(reference.Culture) ? "neutral" : reference.Culture;
			}
		}

		private bool HasReferenceTo(Project project, string projectName)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var refs = new System.Collections.Generic.Dictionary<string, string>();
			if (project.Object is VSLangProj.VSProject vsProject)
			{
				foreach (VSLangProj.Reference reference in vsProject.References)
				{
					var identity = reference.Identity;
					var name = reference.Name;
					if (identity == projectName || name == projectName)
					{
						return true;
					}
				}
			}

			return false;
		}

		private static System.Collections.Generic.List<Project> GetEnvDteProjects(DTE appObject)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var projects = new System.Collections.Generic.List<EnvDTE.Project>(appObject.Solution.Projects.Count);

			foreach (var proj in appObject.Solution.Projects)
			{
				var project = (EnvDTE.Project)proj;

				if (!string.IsNullOrWhiteSpace(project.FullName))
				{
					projects.Add(project);
				}
			}

			return projects;
		}
	}
}
