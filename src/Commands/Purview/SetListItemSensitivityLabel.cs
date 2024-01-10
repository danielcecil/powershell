using Microsoft.SharePoint.Client;
using PnP.PowerShell.Commands.Base.PipeBinds;
using PnP.PowerShell.Commands.Enums;
using PnP.PowerShell.Commands.Model.Graph.Files;
using PnP.PowerShell.Commands.Utilities.REST;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PnP.PowerShell.Commands.Purview
{
    [Cmdlet(VerbsCommon.Set, "PnPListItemSensitivityLabel")]
    [OutputType(typeof(void))]
    public class SetListItemSensitivityLabel : PnPWebCmdlet
    {
        const string ParameterSet_SET = "Set the Sensitivity Label";
        const string ParameterSet_CLEAR = "Clear the Sensitivity Label";

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, ParameterSetName = ParameterSet_SET)]
        [Parameter(ParameterSetName = ParameterSet_CLEAR)]
        public ListItemPipeBind Identity;

        [Parameter(Mandatory = false, ParameterSetName = ParameterSet_SET)]
        [Parameter(ParameterSetName = ParameterSet_CLEAR)]
        public ListPipeBind List;

        [Parameter(Mandatory = true, ParameterSetName = ParameterSet_SET)]
        public SensitivityLabelPipeBind Label;

        [Parameter(Mandatory = false, ParameterSetName = ParameterSet_SET)]
        [Parameter(ParameterSetName = ParameterSet_CLEAR)]
        public SensitivityLabelPipeBind PreviousLabel;

        [Parameter(Mandatory = false, ParameterSetName = ParameterSet_SET)]
        public string JustificationText = String.Empty;

        [Parameter(Mandatory = false, ParameterSetName = ParameterSet_SET)]
        [Parameter(ParameterSetName = ParameterSet_CLEAR)]
        public SensitivityLabelAssignmentMethod AssignmentMethod = SensitivityLabelAssignmentMethod.Privileged;

        [Parameter(Mandatory = false, ParameterSetName = ParameterSet_CLEAR)]
        public SwitchParameter ClearLabel;

        protected override void ExecuteCmdlet()
        {
            List list;
            if (List != null)
            {
                list = List.GetList(CurrentWeb, l => l.Id, l => l.Title);
            }
            else
            {
                if (Identity.Item == null)
                {
                    throw new PSArgumentException($"No -List has been provided specifying the list to update the item in", nameof(List));
                }

                list = Identity.Item.ParentList;
                list.Context.Load(list, l => l.Id, l => l.Title);
                list.Context.ExecuteQueryRetry();
            }

            if (Identity == null || (Identity.Item == null && Identity.Id == 0))
            {
                throw new PSArgumentException($"No -Identity has been provided specifying the item to update", nameof(Identity));
            }

            ListItem item = Identity.GetListItem(list)
                ?? throw new PSArgumentException($"Provided -Identity is not valid.", nameof(Identity)); ;

            string labelId;

            if (ClearLabel.IsPresent)
            {
                labelId = String.Empty;
            }
            else
            {
                if (ParameterSpecified(nameof(Label)) && Label.LabelId == null)
                {
                    var labelLookup = Label.GetLabelByNameThroughGraph(Connection, GraphAccessToken);
                    if (labelLookup == null)
                    {
                        throw new PSArgumentException($"Provided -Label is not valid. Try passing in a Label or Id from the Get-PnPAvailableSensitivityLabel command.", nameof(Label));
                    }
                    labelId = labelLookup.Id.ToString();
                }
                else
                {
                    labelId = Label.LabelId.ToString();
                }
            }

            string prevLabelId = ParameterSpecified(nameof(PreviousLabel)) ? PreviousLabel.LabelId.ToString() : String.Empty;
            if (ParameterSpecified(nameof(PreviousLabel)) && PreviousLabel.LabelId == null)
            {
                var prevLabelLookup = PreviousLabel.GetLabelByNameThroughGraph(Connection, GraphAccessToken);
                if (prevLabelLookup == null)
                {
                    throw new PSArgumentException($"Provided -PreviousLabel is not valid. Try passing in a Label or Id from the Get-PnPAvailableSensitivityLabel command.", nameof(Label));
                }
                prevLabelId = prevLabelLookup.Id.ToString();
            }

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

            string url = "/v2.1/drives/" + listDrive.Id + "/items/" + itemUid + "/setsensitivityLabel";

            string content = JsonSerializer.Serialize(new
            {
                id = labelId,
                assignmentMethod = AssignmentMethod.ToString(),
                justificationText = JustificationText,
                ifMatchLabelId = prevLabelId
            });

            var setLabelResponse = RestHelper.ExecutePostRequest(ClientContext, url, content);
        }
    }
}