using qaImageViewer.Models;
using qaImageViewer.Repository;
using qaImageViewer.Service;
using qaImageViewer.Tasks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Managers
{
    class TaskManager
    {
        static public async Task<int> Launch(ConnectionManager cm, TaskInterface task, IProgress<int> progress)
        {
            bool provideUpdates = ConfigRepository.GetBooleanOption(cm, "TaskManager.ProvideUpdates", true);
            CallBack callback = ()=> { };

            AppTask t = new AppTask {
                Status = AppTaskStatus.EXECUTING,
                Type = task.GetType().FullName,
                Data = task.GetTaskData()
            };


            t.Id = TaskRepository.InsertTask(cm, t);

            task.TaskId = t.Id;

            if (provideUpdates)
            {
                callback = () => { TaskRepository.UpdateTask(cm, t); };
            }

            try
            {
                await Task.Run(() => { task.Execute(progress, callback); });
                t.Status = AppTaskStatus.SUCCESS;
                TaskRepository.UpdateTask(cm, t);
                return t.Id;
            } catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                ProcessingExceptionRepository.InsertProcessingException(cm, new ProcessingExceptionListItem
                {
                    TaskId = t.Id,
                    RowIndex = -1,
                    Type = t.Type,
                    ErrorTrace = ex.ToString(),
                    ResultSetId = -1,
                });
                t.Status = AppTaskStatus.ERROR;
                TaskRepository.UpdateTask(cm, t);
                throw new TaskException(ex.Message);
            }
        }

    }
}
