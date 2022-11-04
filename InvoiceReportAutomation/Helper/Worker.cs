using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Timers;
using System.Threading.Tasks;

namespace InvoiceReportAutomation.Helper
{
    public class Worker
    {
        public readonly Timer _timer;

        public Worker()
        {
            _timer = new Timer(3000) { AutoReset = true };
            _timer.Elapsed += _timer_Elapsed;
        }

        private void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                RequestHelper.ExecuteRequest();
            }
            catch (Exception excep)
            {
                Program.filelogger.Log(excep.ToString());
            }
        }

        public void Start()
        {
            Program.filelogger.Log($"-------Service Start-------");
            _timer.Start();
        }

        public void Stop()
        {
            _timer.Stop();
            Program.filelogger.Log($"-------Service Stop-------");
        }
    }
}
