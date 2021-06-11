using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{

    enum AppTaskStatus
    {
        EXECUTING,
        SUCCESS,
        ERROR,
        None
    }

    class AppTask
    {
        /*
         * id	INTEGER NOT NULL,
                    type        TEXT NOT NULL,
                    start_time INTEGER NOT NULL,
                    update_time INTEGER NOT NULL,
                    data TEXT,
                	status INTEGER NOT NULL,
         */

        public int Id { get; set; }
        public string Type { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime UpdateTime { get; set; }
        public string Data { get; set; }
        public AppTaskStatus Status { get; set; }
    }
}
