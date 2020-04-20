using Polly;
using System;

namespace NetDemo.Polly
{
    /// <summary>
    /// Polly是一种.NET弹性和瞬态故障处理库，允许我们以非常顺畅和线程安全的方式来执诸如行重试，断路，超时，故障恢复等策略。
    /// 源码地址：https://github.com/App-vNext/Polly
    /// 该库实现了七种恢复策略：
    /// 1. 重试策略（Retry）
    /// 2. 断路器（Circuit-breaker）
    /// 3. 超时（Timeout）
    /// 4. 隔板隔离（Bulkhead Isolation）
    /// 5. 缓存（Cache）
    /// 6. 回退（Fallback）
    /// 7. 策略包装（PolicyWrap）
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Test3();
            Console.WriteLine("Complete!");
        }

        #region 默认策略，重试机制
        /// <summary>
        /// 演示尝试除以0并用Polly指定该异常并重试三次。
        /// </summary>
        public static void Test1()
        {
            try
            {
                var retryTwoTimesPolicy = Policy
                    .Handle<DivideByZeroException>()
                    .Retry(3, (ex, count) =>
                    {
                        Console.WriteLine("执行失败! 重试次数 {0}", count);
                        Console.WriteLine("异常来自 {0}", ex.GetType().Name);
                    });

                retryTwoTimesPolicy.Execute(() =>
                {
                    Compute();
                });
            }
            catch (DivideByZeroException e)
            {
                Console.WriteLine($"Excuted Failed,Message: ({e.Message})");
            }
        }

        static int Compute()
        {
            var a = 0;
            return 1 / a;
        }
        #endregion

        #region 重试策略，重试机制
        /// <summary>
        /// 演示隔一段时间重试调用一次，重复调用几次后仍失败则不再回调
        /// </summary>
        public static void Test2()
        {
            try
            {
                var politicaWaitAndRetry = Policy
                    .Handle<DivideByZeroException>()
                    .WaitAndRetry(new[]
                    {
                        TimeSpan.FromSeconds(1),
                        TimeSpan.FromSeconds(3),
                        TimeSpan.FromSeconds(5),
                        TimeSpan.FromSeconds(7)
                    }, ReportaError);

                politicaWaitAndRetry.Execute(() =>
                {
                    ZeroExcepcion();
                });
            }
            catch (Exception e)
            {
                Console.WriteLine($"Executed Failed,Message:({e.Message})");
            }
        }
        static void ZeroExcepcion()
        {
            throw new DivideByZeroException();
        }

        static void ReportaError(Exception e, TimeSpan tiempo, int intento, Context contexto)
        {
            Console.WriteLine($"异常: {intento:00} (调用秒数: {tiempo.Seconds} 秒)\t执行时间: {DateTime.Now}");
        }
        #endregion

        #region 反馈策略
        /// <summary>
        /// 执行失败后返回的结果，此时要为Polly指定返回类型，然后指定异常，最后调用Fallback方法
        /// </summary>
        public static void Test3()
        {
            var fallBackPolicy =
                Policy<string>
                    .Handle<Exception>()
                    .Fallback("执行失败，返回Fallback");

            var fallBack = fallBackPolicy.Execute(() =>
            {
                return ThrowException();
            });

            Console.WriteLine(fallBack);
        }

        static string ThrowException()
        {
            throw new Exception();
        }
        #endregion
    }
}
