using System;
using System.Collections.Generic;
using System.Linq;

namespace NetDemo.EFEvent
{
    public interface IEventPublisher
    {
        void Publish<T>(T eventMessage);
    }

    public class EventPublisher : IEventPublisher
    {
        #region Fields

        private readonly ISubscriptionService _subscriptionService;

        #endregion

        #region Ctor

        public EventPublisher(ISubscriptionService subscriptionService)
        {
            _subscriptionService = subscriptionService;
        }

        #endregion

        #region Utilities
        protected virtual void PublishToConsumer<T>(IConsumer<T> x, T eventMessage)
        {
            try
            {
                x.HandleEvent(eventMessage);
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
            }
        }

        #endregion

        #region Methods
        public virtual void Publish<T>(T eventMessage)
        {
            //get all event subscribers, excluding from not installed plugins
            //var subscribers = _subscriptionService.GetSubscriptions<T>()
            //    .Where(subscriber => PluginManager.FindPlugin(subscriber.GetType())?.Installed ?? true).ToList();

            var subscribers = _subscriptionService.GetSubscriptions<T>().ToList();
            //publish event to subscribers
            subscribers.ForEach(subscriber => PublishToConsumer(subscriber, eventMessage));
        }

        #endregion
    }
}
