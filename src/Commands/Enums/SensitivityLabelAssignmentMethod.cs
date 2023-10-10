namespace PnP.PowerShell.Commands.Enums
{
    /// <summary>
    /// Enum that defines the possible types of users
    /// </summary>
public enum SensitivityLabelAssignmentMethod
    {
        /// <summary>
        /// The assignment method for the label is standard.
        /// </summary>
        standard,
        
        /// <summary>
        /// The assignment method for the label is privileged. Indicates that the label is applied manually by a user or by an admin.
        /// </summary>
        privileged,

        /// <summary>
        ///  Indicates that the label is applied automatically by the system due to a configured policy, such as default label or autoclassification of sensitive content.
        /// </summary>
        auto
    }
}
