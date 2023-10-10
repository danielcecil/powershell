using System;
using System.Text.Json.Serialization;

namespace PnP.PowerShell.Commands.Model.Graph.Purview
{
    public class ExtractSensitivityLabelsResult
    {
        /// <summary>
        /// The label ID is a globally unique identifier (GUID)
        /// </summary>
        [JsonPropertyName("labels")]
        public SensitivityLabelAssignment[] Labels { get; set; }

    }
}