using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text.Json.Serialization;

namespace PnP.PowerShell.Commands.Model.Graph.Files
{
    public class DriveResult
    {
        [JsonPropertyName("value")]
        public List<Drive> Drives { get; set; }
    }
}