using System;
using System.IO;

namespace MsgViewerWeb
{
    public class Global : System.Web.HttpApplication
    {

        void Application_Start(object sender, EventArgs e)
        {
            // Code that runs on application startup
            CleanTempFolders();
        }

        void Application_End(object sender, EventArgs e)
        {
            //  Code that runs on application shutdown
            CleanTempFolders();
        }

        void Application_Error(object sender, EventArgs e)
        {
            // Code that runs when an un-handled error occurs
        }

        void Session_Start(object sender, EventArgs e)
        {
            // Code that runs when a new session is started
        }

        void Session_End(object sender, EventArgs e)
        {
            // Code that runs when a session ends. 
            // Note: The Session_End event is raised only when the session-state mode
            // is set to InProc in the Web.config file. If session mode is set to StateServer 
            // or SQLServer, the event is not raised.            
        }

        private void CleanTempFolders()
        {
            var path = Server.MapPath("~/FileSystem/Messages");

            try
            {
                foreach (var dir in Directory.GetDirectories(path))
                {
                    try
                    {
                        Directory.Delete(dir, true);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                foreach (var file in Directory.GetFiles(path))
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
