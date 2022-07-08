using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace FastExcelDNA.ExcelDNA.RTDServer
{
    // Helper Class for Queueing the tasks and do work
    public class ThreadPoolQueue
    {
        private readonly object sync = new object();
        private readonly LinkedList<Action> queue = new LinkedList<Action>();
        private bool isActive;

        public void Enqueue(Action work)
        {
            var node = new LinkedListNode<Action>(work);
            lock (sync)
            {
                queue.AddLast(node);
                if (isActive)
                {
                    return;
                }
                isActive = true;
            }
            ThreadPool.QueueUserWorkItem(_ => ProcessQueue());
        }

        private void ProcessQueue()
        {
            while (true)
            {
                var action = GetNext();
                if (action == null)
                {
                    break;
                }
                action();
            }
        }

        private Action GetNext()
        {
            lock (sync)
            {
                if (queue.Count == 0)
                {
                    isActive = false;
                    return null;
                }
                var first = queue.First;
                queue.RemoveFirst();
                return first.Value;
            }
        }
    }
}
