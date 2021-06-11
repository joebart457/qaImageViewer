using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Tasks
{
    class TaskException: Exception
    {
        private string _message { get; set; }

        public TaskException(string msg)
        {
            this._message = msg;
        }

        public override string ToString()
        {
            return "TaskException:" + _message;
        }
    }


    public delegate void CallBack();

    interface TaskInterface
    {
        public int TaskId { get; set; }
        void Execute(IProgress<int> progress, CallBack callback);
        string GetTaskData();
    }
}
