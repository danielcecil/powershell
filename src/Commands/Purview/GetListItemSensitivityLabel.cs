using Microsoft.SharePoint.Client;
using PnP.PowerShell.Commands.Base.PipeBinds;
using PnP.PowerShell.Commands.Model.Graph.Purview;
using System;
using System.Management.Automation;

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


            object labelGuid;
            item.FieldValues.TryGetValue("_IpLabelId", out labelGuid);

            if (!String.IsNullOrEmpty((string)labelGuid))
            {
                WriteObject(new SensitivityLabelPipeBind((string)labelGuid).GetLabelByIdThroughGraph(Connection, GraphAccessToken));
            }

        }
    }
}