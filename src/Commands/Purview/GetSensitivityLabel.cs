using Microsoft.SharePoint.Client;
using PnP.PowerShell.Commands.Base.PipeBinds;
using PnP.PowerShell.Commands.Model.Graph.Purview;
using System;
using System.Management.Automation;

namespace PnP.PowerShell.Commands.Purview
{
    [Cmdlet(VerbsCommon.Get, "PnPSensitivityLabel")]
    [OutputType(typeof(InformationProtectionLabel))]
    public class GetSensitivityLabel : PnPWebCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListItemPipeBind ListItem;
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
                if (ListItem.Item == null)
                {
                    throw new PSArgumentException($"No -List has been provided specifying the list to update the item in", nameof(ListItem));
                }

                list = ListItem.Item.ParentList;
            }

            if (ListItem == null || (ListItem.Item == null && ListItem.Id == 0))
            {
                throw new PSArgumentException($"No -ListItem has been provided specifying the item to update", nameof(ListItem));
            }

            ListItem item = ListItem.GetListItem(list)
                ?? throw new PSArgumentException($"Provided -ListItem is not valid.", nameof(ListItem)); ;


            object labelGuid;
            item.FieldValues.TryGetValue("_IpLabelId", out labelGuid);

            if (!String.IsNullOrEmpty((string)labelGuid))
            {
                WriteObject(new SensitivityLabelPipeBind((string)labelGuid).GetLabelByIdThroughGraph(Connection, GraphAccessToken));
            }

        }
    }
}