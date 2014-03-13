using System;
using System.IO;
using System.Linq;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;

namespace MsgViewerWeb
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Init(object sender, EventArgs e)
        {

        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                foreach (var fileName in Directory.EnumerateFiles(Server.MapPath("~/FileSystem"), "*.msg").Select(Path.GetFileName))
                {                    
                    var a = new HtmlAnchor
                        {
                            HRef = "MsgHandler.ashx?file=" + fileName,
                            InnerText = fileName,
                            CausesValidation = false
                        };
                    messages.Controls.Add(a);
                }
            }
        }
    }
}
