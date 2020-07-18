using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  public sealed class EventfulConcurrentQueue<T>
  {
    public readonly AutoResetEvent DataAvailableEvent = new AutoResetEvent(false);

    private ConcurrentQueue<T> _queue;

    public EventfulConcurrentQueue()
    {
      _queue = new ConcurrentQueue<T>();
    }

    public void Enqueue(T item)
    {
      _queue.Enqueue(item);
      DataAvailableEvent.Set();
    }

    public bool TryDequeue(out T result)
    {
      return _queue.TryDequeue(out result);
    }
  }
}
