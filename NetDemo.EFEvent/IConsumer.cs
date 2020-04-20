using System;
using System.Collections.Generic;
using System.Text;

namespace NetDemo.EFEvent
{
    /// <summary>
    /// Consumer interface
    /// </summary>
    /// <typeparam name="T">Type</typeparam>
    public interface IConsumer<T>
    {
        /// <summary>
        /// Handle event
        /// </summary>
        /// <param name="eventMessage">Event</param>
        void HandleEvent(T eventMessage);
    }
}
