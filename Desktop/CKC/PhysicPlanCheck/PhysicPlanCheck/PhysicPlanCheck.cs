using PhysicPlanCheck;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
//using Syncfusion.XlsIO;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
// [assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{ 
    public class Script
    {
        public Script()
        {
        }

        public static ScriptContext context { get; set; }


        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context_input, System.Windows.Window window, ScriptEnvironment environment)
        {
            // TODO : Add here the code that is called when the script is launched from Eclipse.
            //Main(context_input);

            ViewModel vm = new ViewModel(context_input);
            UserControl1 control = new UserControl1();
            control.DataContext = vm;
            window.Content = control;

            //yitingEdit
            window.Height = 810;
            window.Width = 660;
        }
    }
}
