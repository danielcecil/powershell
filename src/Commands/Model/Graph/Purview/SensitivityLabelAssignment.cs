using System;
using System.Text.Json.Serialization;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace PnP.PowerShell.Commands.Model.Graph.Purview
{
    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
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
        public string SensitivityLabelAssignmentMethod { get; set; }

        [JsonPropertyName("justificationText")]
        public string JustificationText { get; set; }

    }
}