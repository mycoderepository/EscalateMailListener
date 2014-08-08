using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace EscalateService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        private System.Timers.Timer _timer;
        protected override void OnStart(string[] args)
        {
            // TODO: Add code here to start your service.
#if DEBUG
            System.Diagnostics.Debugger.Launch(); // This will automatically prompt to attach the debugger if you are in Debug configuration
#endif

            _timer = new System.Timers.Timer(10 * 1000); //10 seconds
            _timer.Elapsed += this.TimerOnElapsed;
            _timer.Start();

        }
        private void TimerOnElapsed(object sender, ElapsedEventArgs elapsedEventArgs)
        {
            // Call to run off to a database or do some processing
            ProcessStartInfo info = new ProcessStartInfo(@"C:\ETML\Escaltethreshold.exe");
            info.UseShellExecute = false;
            info.RedirectStandardError = true;
            info.RedirectStandardInput = true;
            info.RedirectStandardOutput = true;
            info.CreateNoWindow = true;
            info.ErrorDialog = false;
            info.WindowStyle = ProcessWindowStyle.Hidden;

            Process process = Process.Start(info);
        }

        protected override void OnStop()
        {
            // TODO: Add code here to perform any tear-down necessary to stop your service.
            _timer.Stop();
            _timer.Elapsed -= TimerOnElapsed;
        }
    }
}
