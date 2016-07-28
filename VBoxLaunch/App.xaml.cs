using System.Windows;
using System.Windows.Shell;
using VirtualBox;

namespace VBoxLaunch
{
	/// <summary>
	/// Interaction logic for App.xaml
	/// </summary>
	public partial class App : Application
	{
		protected override void OnStartup(StartupEventArgs e)
		{
			base.OnStartup(e);

			VirtualBox.IVirtualBoxClient vBox = new VirtualBox.VirtualBoxClient();

			if (e.Args.Length > 0)
			{
				VirtualBox.IMachine machine = vBox.VirtualBox.FindMachine(e.Args[0]);

				if (machine != null)
				{
					if (machine.State == MachineState.MachineState_Paused || machine.State == MachineState.MachineState_PoweredOff || machine.State == MachineState.MachineState_Saved || machine.State == MachineState.MachineState_Aborted)
					{
						Session session = vBox.Session;

						IProgress progress = machine.LaunchVMProcess(session, string.Empty, string.Empty);
						progress.WaitForCompletion(-1);
					}
				}
			}

			JumpList jumpList = JumpList.GetJumpList(this);
			if (jumpList == null)
			{
				jumpList = new JumpList();
			}

			jumpList.ShowFrequentCategory = false;
			jumpList.ShowRecentCategory = false;

			for (int i = jumpList.JumpItems.Count - 1; i >= 0; --i)
			{
				if (jumpList.JumpItems[i].GetType() != typeof(JumpTask))
				{
					jumpList.JumpItems.RemoveAt(i);
				}
				else
				{
					bool keep = false;
					foreach (VirtualBox.IMachine machine in vBox.VirtualBox.Machines)
					{
						keep = (machine.Name == ((JumpTask)jumpList.JumpItems[i]).Title && machine.Name == ((JumpTask)jumpList.JumpItems[i]).Arguments);
						if (keep)
						{
							break;
						}
					}
					if (!keep)
					{
						jumpList.JumpItems.RemoveAt(i);
					}
				}
			}

			for (int i = 0; i < vBox.VirtualBox.Machines.Length; ++i)
			{
				VirtualBox.IMachine machine = (VirtualBox.IMachine)vBox.VirtualBox.Machines.GetValue(i);
				bool add = true;
				foreach (JumpTask task in jumpList.JumpItems)
				{
					add = (machine.Name != task.Title || machine.Name != task.Arguments);
					if (!add)
					{
						break;
					}
				}
				if (add)
				{
					JumpTask task = new JumpTask
					{
						Title = machine.Name,
						Arguments = machine.Name
					};

					task.CustomCategory = "Machines";

					jumpList.JumpItems.Add(task);
				}
			}

			JumpList.SetJumpList(Application.Current, jumpList);

			Application.Current.Shutdown();
		}
	}
}
