using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;

namespace CustomActions
{
    public class CustomActions
    {
        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiGetProductInfo(string product, string property, [Out] StringBuilder valueBuf, ref int len);

        [CustomAction]
        public static ActionResult ExtractPreviousVersion(Session session)
        {
            string productId = session["PREVIOUSFOUND"];
            if (!string.IsNullOrEmpty(productId))
            {
                int length = 32;
                StringBuilder sb = new StringBuilder();
                if (0 == MsiGetProductInfo(productId, "VersionString", sb, ref length))
                {
                    session.Log("ExtractPreviousVersion: " + sb);
                }
            }

            return ActionResult.Success;
        }
    }
}
