using log4net;
using log4net.Config;
using System.IO;

namespace ProjectImports
{
    public static class Log<T> where T : class
    {
        public static ILog Logger
        {
            get { return LogManager.GetLogger(typeof(T)); }
        }
        public static void Configure(string pathToConfigFile)
        {
            var fileInfo = new FileInfo(pathToConfigFile);
            XmlConfigurator.ConfigureAndWatch(fileInfo);
        }
    }
}
