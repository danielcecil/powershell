using System;
using System.Text.Json.Serialization;

namespace PnP.PowerShell.Commands.Model.Graph.Purview
{
    public class SensitivityLabelAssignment
    {
        /// <summary>
        /// The label ID is a globally unique identifier (GUID)
        /// </summary>
        [JsonPropertyName("sensitivityLabelId")]
        public Guid? SensitivityLabelId { get; set; }

        [JsonPropertyName("tenantId")]
        public Guid? TenantId { get; set; }

        [JsonPropertyName("assignmentMethod")]
        public Enums.SensitivityLabelAssignmentMethod SensitivityLabelAssignmentMethod { get; set; }

        [JsonPropertyName("justificationText")]
        public string JustificationText { get; set; }

    }
}