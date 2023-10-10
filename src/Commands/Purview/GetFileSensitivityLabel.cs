using Microsoft.IdentityModel.Tokens;
using Microsoft.SharePoint.Client;
using PnP.PowerShell.Commands.Base;
using PnP.PowerShell.Commands.Utilities.REST;
using System.Management.Automation;
using System.Net.Http;
using PnP.PowerShell.Commands.Model.Graph.Purview;

namespace PnP.PowerShell.Commands.Purview
{
    [Cmdlet(VerbsCommon.Get, "PnPFileSensitivityLabel")]
    [OutputType(typeof(SensitivityLabelAssignment[]))]
    public class GetFileSensitivityLabel : PnPGraphCmdlet
    {
        [Parameter(Mandatory = true)]
        public string Path = string.Empty;
        protected override void ExecuteCmdlet()
        {

            if (Path.IsNullOrEmpty() == true)
            {
                WriteWarning("The file Path is not specified");
                return;
            }

            string base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(Path));
            string encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');

            string url = $"/beta/shares/{encodedUrl}/driveItem/extractSensitivityLabels";

            var stringContent = new StringContent("{}");
            stringContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
            var response = GraphHelper.PostAsync<ExtractSensitivityLabelsResult>(Connection, url, stringContent, AccessToken).GetAwaiter().GetResult();

        if(response == null) {
            return;
        }
            WriteObject(response.Labels, true);
        }
    }
}