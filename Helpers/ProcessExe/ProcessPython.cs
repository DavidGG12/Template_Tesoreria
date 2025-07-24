using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.Helpers.ProcessExe
{
    public class ProcessPython
    {
        private string _pathExe;
        private ProcessStartInfo _start;

        public ProcessPython(string pathExe)
        {
            this._pathExe = pathExe;
            this._start = new ProcessStartInfo();
            this._start.FileName = this._pathExe;
            this._start.UseShellExecute = false;
            this._start.RedirectStandardOutput = true;
            this._start.RedirectStandardError = true;
            this._start.CreateNoWindow = true;
        }

       public string ExecuteProcess()
       {
            using(Process process = Process.Start(this._start))
            {
                var output = process.StandardOutput.ReadToEnd();    
                var error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if(!string.IsNullOrEmpty(error))
                    return error;

                return output;
            }
       }
    }
}
