using Microsoft.IdentityModel.Tokens;
using Microsoft.SharePoint.Client;
using PnP.PowerShell.Commands.Base;
using PnP.PowerShell.Commands.Utilities;
using PnP.PowerShell.Commands.Utilities.REST;
using System.Management.Automation;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System;
using PnP.PowerShell.Commands.Model.Graph.Purview;
using PnP.PowerShell.Commands.Enums;

namespace PnP.PowerShell.Commands.Purview
{
    [Cmdlet(VerbsCommon.Set, "PnPFileSensitivityLabel")]
    // [OutputType(typeof(PnP.PowerShell.Commands.Model.SharePoint.SensitivityLabel))]
    public class SetFileSensitivityLabel : PnPGraphCmdlet
    {
        [Parameter(Mandatory = true)]
        public string Path = string.Empty;

        [Parameter(Mandatory = false)]
        public InformationProtectionLabel Label;

        [Parameter(Mandatory = false)]
        public Guid? LabelId;

        [Parameter(Mandatory = false)]
        public string JustificationText = string.Empty;

        // More info: https://learn.microsoft.com/en-us/graph/api/resources/sensitivitylabelassignment
        public const string ASSIGNMENT_METHOD = "privileged";
        protected override void ExecuteCmdlet()
        {

            if (Path.IsNullOrEmpty() == true)
            {
                WriteWarning("The file Path is not specified");
                return;
            }

            if (Label == null && LabelId == Guid.Empty)
            {
                WriteWarning("The Label or LabelId is not specified");
                return;
            }

            Guid? sensitivityLabelId;
            if (Label != null)
            {
                sensitivityLabelId = Label.Id;
            }
            else
            {
                sensitivityLabelId = LabelId;
            }


            string base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(Path));
            string encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
            string url = $"/beta/shares/{encodedUrl}/driveItem/assignSensitivityLabel";

            var jsonSerializer = new JsonSerializerOptions { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull, PropertyNamingPolicy = JsonNamingPolicy.CamelCase};
            jsonSerializer.Converters.Add(new JsonStringEnumConverter());

            string json = JsonSerializer.Serialize(new SensitivityLabelAssignment { SensitivityLabelId = sensitivityLabelId, SensitivityLabelAssignmentMethod = "privileged", JustificationText = JustificationText }, jsonSerializer);
            var stringContent = new StringContent(json);
            stringContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
            var response = GraphHelper.PostAsync(Connection, url, stringContent, AccessToken).GetAwaiter().GetResult();

            if (response == null)
            {
                return;
            }

            WriteObject(response, false);
        }
    }
}