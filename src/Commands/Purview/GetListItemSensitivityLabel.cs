using Microsoft.SharePoint.Client;
using PnP.PowerShell.Commands.Base.PipeBinds;
using PnP.PowerShell.Commands.Enums;
using PnP.PowerShell.Commands.Model.Graph.Files;
using PnP.PowerShell.Commands.Model.Graph.Purview;
using PnP.PowerShell.Commands.Utilities.REST;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection.Emit;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PnP.PowerShell.Commands.Purview
{
    [Cmdlet(VerbsCommon.Get, "PnPListItemSensitivityLabel")]
    [OutputType(typeof(InformationProtectionLabel))]
    public class GetListItemSensitivityLabel : PnPWebCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListItemPipeBind Identity;
        [Parameter(Mandatory = false)]
        public ListPipeBind List;

        protected override void ExecuteCmdlet()
        {
            List list;
            if (List != null)
            {
                list = List.GetList(CurrentWeb);
            }
            else
            {
                if (Identity.Item == null)
                {
                    throw new PSArgumentException($"No -List has been provided specifying the list to update the item in", nameof(Identity));
                }

                list = Identity.Item.ParentList;
            }

            if (Identity == null || (Identity.Item == null && Identity.Id == 0))
            {
                throw new PSArgumentException($"No -Identity has been provided specifying the item to update", nameof(Identity));
            }

            ListItem item = Identity.GetListItem(list)
                ?? throw new PSArgumentException($"Provided -Identity is not valid.", nameof(Identity)); ;

            Guid listId = list.Id;
            string listTitle = list.Title;

            // Get the Drive objects for the current site
            string getDrivesResponse = RestHelper.ExecuteGetRequest(ClientContext, "/v2.1/drives");
            DriveResult driveResult = JsonSerializer.Deserialize<DriveResult>(getDrivesResponse);

            // Find the Drive object for the current list
            Drive listDrive = driveResult.Drives.FirstOrDefault(i => i.Name == listTitle);

            if (listDrive == null)
            {
                throw new PSArgumentException($"The -List that has been provided could not be found. Check that the List is a Document Library.", nameof(List));
            }

            // Ensure the UniqueId is present on the ListItem object.
            object itemUid;
            if (!item.FieldValues.TryGetValue("UniqueId", out itemUid))
            {
                throw new PSArgumentNullException($"The -Identity does not contain a UniqueId property.", nameof(Identity));
            }

            string url = "/v2.1/drives/" + listDrive.Id + "/items/" + itemUid + "/extractSensitivityLabels";

            ExtractSensitivityLabelsResult extractLabelResponse = RestHelper.ExecutePostRequest<ExtractSensitivityLabelsResult>(ClientContext, url, String.Empty);

            if (extractLabelResponse.Labels.Count > 0)
            {
                // TODO: Check for multiple Labels and fetch the one with a matching TenantId or First()
                WriteObject(new SensitivityLabelPipeBind(extractLabelResponse.Labels.First().LabelId).GetLabelByIdThroughGraph(Connection, GraphAccessToken));
            }
        }
    }
    public class ExtractSensitivityLabelsResult
    {
        [JsonPropertyName("labels")]
        public List<SensitivityLabelAssignment> Labels { get; set; }
    }

    public class SensitivityLabelAssignment
    {
        [JsonPropertyName("assignmentMethod")]
        public string AssignmentMethod { get; set; }

        [JsonPropertyName("sensitivityLabelId")]
        public string LabelId { get; set; }

        [JsonPropertyName("tenantId")]
        public string TenantId { get; set; }
    }
}