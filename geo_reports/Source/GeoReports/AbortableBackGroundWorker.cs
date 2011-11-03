using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Threading;

namespace GeoReports
{
    public class AbortableBackgroundWorker : BackgroundWorker
    {

        private Thread workerThread;
        private int progress = 0;
        private Boolean isAlive;
        public int PercentProgress
        {
            get
            {
                return progress;
            }
            set
            {
                progress = value;
            }
        }
        public bool IsAlive
        {
            get
            {
                return isAlive;
            }
        }
        protected override void OnDoWork(DoWorkEventArgs e)
        {

            isAlive = true;
            workerThread = Thread.CurrentThread;
            try
            {
                base.OnDoWork(e);
            }
            catch (ThreadAbortException)
            {
                e.Cancel = true;
                Thread.ResetAbort();
            }
        }
        protected override void OnProgressChanged(ProgressChangedEventArgs args)
        {

            progress = args.ProgressPercentage;
            base.OnProgressChanged(args);

        }
        public Thread GetCurrentThred()
        {
            return workerThread;
        }
        protected override void OnRunWorkerCompleted(RunWorkerCompletedEventArgs e)
        {

            isAlive = false;
            base.OnRunWorkerCompleted(e);

        }
        public void Abort()
        {

            isAlive = false;
            if (workerThread != null)
            {
                workerThread.Abort();
                workerThread = null;
            }
        }
    }
    

}
