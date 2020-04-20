using Autofac;
using Autofac.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace NetDemo.EFEvent
{
    public class NopEngine : IEngine
    {
        private IServiceProvider _serviceProvider { get; set; }

        public IServiceProvider ConfigureServices()
        {
            var containerBuilder = new ContainerBuilder();

            containerBuilder.RegisterType<EventPublisher>().As<IEventPublisher>();
            containerBuilder.RegisterType<SubscriptionService>().As<ISubscriptionService>();

            //方法一
            containerBuilder.RegisterType<CacheEventConsumer>().As<IConsumer<EntityInsertedEvent<BlogPost>>>();

            //方法二
            var consumer = typeof(MailEventConsumer);
            containerBuilder.RegisterType(consumer).As(consumer.FindInterfaces((type, criteria) =>
            {
                return true;
            }, typeof(IConsumer<>)));

            //create service provider
            _serviceProvider = new AutofacServiceProvider(containerBuilder.Build());
            return _serviceProvider;
        }

        public T Resolve<T>() where T : class
        {
            return (T)_serviceProvider.GetRequiredService(typeof(T));
        }

        public object Resolve(Type type)
        {
            return _serviceProvider.GetRequiredService(type);
        }

        public IEnumerable<T> ResolveAll<T>()
        {
            return (IEnumerable<T>)_serviceProvider.GetServices(typeof(T));
        }
    }

    public interface IEngine
    {
        IServiceProvider ConfigureServices();

        /// <summary>
        /// Resolve dependency
        /// </summary>
        /// <typeparam name="T">Type of resolved service</typeparam>
        /// <returns>Resolved service</returns>
        T Resolve<T>() where T : class;

        /// <summary>
        /// Resolve dependency
        /// </summary>
        /// <param name="type">Type of resolved service</param>
        /// <returns>Resolved service</returns>
        object Resolve(Type type);

        /// <summary>
        /// Resolve dependencies
        /// </summary>
        /// <typeparam name="T">Type of resolved services</typeparam>
        /// <returns>Collection of resolved services</returns>
        IEnumerable<T> ResolveAll<T>();
    }

    /// <summary>
    /// Provides access to the singleton instance of the Nop engine.
    /// </summary>
    public class EngineContext
    {
        #region Methods

        /// <summary>
        /// Create a static instance of the Nop engine.
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static IEngine Create()
        {
            //create NopEngine as engine
            return Singleton<IEngine>.Instance ?? (Singleton<IEngine>.Instance = new NopEngine());
        }

        /// <summary>
        /// Sets the static engine instance to the supplied engine. Use this method to supply your own engine implementation.
        /// </summary>
        /// <param name="engine">The engine to use.</param>
        /// <remarks>Only use this method if you know what you're doing.</remarks>
        public static void Replace(IEngine engine)
        {
            Singleton<IEngine>.Instance = engine;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the singleton Nop engine used to access Nop services.
        /// </summary>
        public static IEngine Current
        {
            get
            {
                if (Singleton<IEngine>.Instance == null)
                {
                    Create();
                }

                return Singleton<IEngine>.Instance;
            }
        }

        #endregion
    }
}
