using System;
using System.ComponentModel;
using System.Text.Json.Serialization;

namespace PnP.PowerShell.Commands.Model.Graph.Files
{
    /// <summary>
    /// Describes the Drive object in the Graph API. A Drive is the top-level object that represents a user's OneDrive or a document library in SharePoint.
    /// </summary>
    /// <seealso cref="https://learn.microsoft.com/en-us/graph/api/resources/drive"/>
    public class Drive
    {
        [JsonPropertyName("id")]
        [Description("The unique identifier of the drive. Read-only.")]
        public string Id { get; set; }

        [JsonPropertyName("name")]
        [Description("The name of the item. Read-write.")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        [Description("Provide a user-visible description of the drive. Read-write.")]
        public string Description { get; set; }

        [JsonPropertyName("driveType")]
        [Description("Describes the type of drive represented by this resource. OneDrive personal drives will return personal. OneDrive for Business will return business. SharePoint document libraries will return documentLibrary. Read-only.")]
        public string DriveType { get; set; }

        [JsonPropertyName("createdDateTime")]
        [Description("Date and time of item creation. Read-only.")]
        public DateTime CreatedDateTime { get; set; }

        [JsonPropertyName("lastModifiedDateTime")]
        [Description("Date and time the item was last modified. Read-only.")]
        public DateTime LastModifiedDateTime { get; set; }

        [JsonPropertyName("createdBy")]
        [Description("Identity of the user, device, or application which created the item. Read-only.")]
        public object CreatedBy { get; set; }
        [JsonPropertyName("lastModifiedBy")]
        [Description("Identity of the user, device, and application which last modified the item. Read-only.")]
        public object LastModifiedBy { get; set; }

        [JsonPropertyName("owner")]
        [Description("	Optional. The user account that owns the drive. Read-only.")]
        public object Owner { get; set; }

        [JsonPropertyName("quota")]
        [Description("Optional. Information about the drive's storage space quota. Read-only.")]
        public object Quota { get; set; }

        [JsonPropertyName("sharepointIds")]
        [Description("Returns identifiers useful for SharePoint REST compatibility. Read-only. This property is not returned by default and must be selected using the $select query parameter.")]
        public object SharePointIds { get; set; }

        [JsonPropertyName("system")]
        [Description("If present, indicates that this is a system-managed drive. Read-only.")]
        public object System { get; set; }

        [JsonPropertyName("webUrl")]
        [Description("URL that displays the resource in the browser. Read-only.")]
        public string WebUrl { get; set; }
    }
}