using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.Helpers.MangementLog
{
    public class ExecutionTimer
    {
        private Stopwatch _stopwatch;

        public ExecutionTimer()
        {
            this._stopwatch = new Stopwatch();
        }

        public void startExecution()
        {
            this._stopwatch.Start();
        }

        public string endExecution()
        {
            this._stopwatch.Stop();

            TimeSpan ts = this._stopwatch.Elapsed;

            return $"{ts.Hours:00}:{ts.Minutes:00}:{ts.Seconds:00}.{ts.Milliseconds / 10:00}";
        }
    }
}
