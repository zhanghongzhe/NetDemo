using Autofac;
using Autofac.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NetDemo.EFEvent
{
    public class TestClass
    {
        public static IContainer Container { get; set; }
        public static IServiceProvider _serviceProvider { get; set; }

        public void Start()
        {
            EngineContext.Current.ConfigureServices();

            var blogPost = new BlogPost() { Id = 1, Title = "Title" };

            var eventPublisher = EngineContext.Current.Resolve<IEventPublisher>();
            eventPublisher.EntityInserted(blogPost);
        }
    }

    public partial class CacheEventConsumer : IConsumer<EntityInsertedEvent<BlogPost>>
    {
        public void HandleEvent(EntityInsertedEvent<BlogPost> eventMessage)
        {
            Console.WriteLine("CacheEventConsumer:{0}", eventMessage.Entity.Title);
        }
    }
    public partial class MailEventConsumer : IConsumer<EntityInsertedEvent<BlogPost>>
    {
        public void HandleEvent(EntityInsertedEvent<BlogPost> eventMessage)
        {
            Console.WriteLine("MailEventConsumer:{0}", eventMessage.Entity.Title);
        }
    }

    public abstract partial class BaseEntity
    {
        public int Id { get; set; }
    }

    public partial class BlogPost : BaseEntity
    {
        public string Title { get; set; }
    }
}
