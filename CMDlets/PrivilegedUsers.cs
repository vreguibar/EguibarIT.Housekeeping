using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Function for Privileged Users housekeeping</para>
    /// <para type="description">For each of the user accounts stored within the USERS container, make it part of the corresponding admin tier group. 
    /// For example, if the user account is a Tier2 account(_T2), the user must belong to SG_Tier2Admins group.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-PrivilegedUsersHousekeeping "OU=Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-PrivilegedUsersHousekeeping -AdminUsersDN "OU=Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters and removing Non-Standard users.</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-PrivilegedUsersHousekeeping -AdminUsersDN "OU=Users,OU=Admin,DC=EguibarIT,DC=local" -DisableNonStandardUsers</code>
    ///         </para>
    /// </example>
    /// <remarks>Ensure each Semi-Privileged account is member of the corresponding tier group.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "PrivilegedUsersHousekeeping", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(int))]
    public class PrivilegedUsersHousekeeping : PSCmdlet
    {

        int i = 0;
        int t0 = 0;
        int t1 = 0;
        int t2 = 0;
        readonly int myId = 0;

        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[STRING] Admin User Account OU Distinguished Name. (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        /// <para type="description">Distinguished Name of the container where the Admin users are located.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Distinguished Name of the container where the Admin Accounts are located."
            )]
        [ValidateNotNullOrEmpty]
        public string AdminUsersDN
        {
            get { return _adminusersdn; }
            set { _adminusersdn = value; }
        }

        private string _adminusersdn;

        /// <summary>
        /// <para type="inputType">SWITCH parameter (true or false). If present the value becomes TRUE, and </para>
        /// <para type="description">Switch indicator to disable all Non-Standard users</para>
        /// </summary>
        [Parameter(
               Position = 1,
               Mandatory = false,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Switch indicator. If present (true), will disable all Non-Standard users."
            )]
        public SwitchParameter DisableNonStandardUsers
        {
            get { return _disablenonstandardusers; }
            set { _disablenonstandardusers = value; }
        }

        private bool _disablenonstandardusers;

        #endregion Parameters definition

        #region Begin()

        /// <summary>
        ///
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (this.MyInvocation.BoundParameters.ContainsKey("Verbose"))
            {
                WriteVerbose("|=> ************************************************************************ <=|");
                WriteVerbose(DateTime.Today.ToShortDateString());
                WriteVerbose(string.Format("  Starting: {0}", this.MyInvocation.MyCommand));

                string paramVerbose;
                paramVerbose = "Parameters:\n";
                paramVerbose += string.Format("{0,-12}{1,-30}{2,-30}\n", null, " Key", " Value");
                paramVerbose += string.Format("{0,-12}{1,-30}{2,-30}\n", null, "----------", "----------");

                // display PSBoundparameters formatted nicely for Verbose output
                // var is iDictionary
                var pb = this.MyInvocation.BoundParameters; // | Format - Table - AutoSize | Out - String).TrimEnd()

                foreach (var item in pb)
                {
                    paramVerbose += string.Format("{0,-12}{1,-30}{2,-30}\n", null, item.Key, item.Value);
                }
                WriteVerbose(string.Format("{0}\n", paramVerbose));
            }
        }

        #endregion Begin()

        #region Process()

        /// <summary>
        ///
        /// </summary>
        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            // Set-PrivilegedUsersHousekeeping -AdminUsersDN 'OU=Admin Accounts,OU=Admin,DC=EguibarIT,DC=local' -Verbose

            // Define the Progress Record (Progress Bar to be displayed)            
            string myActivity = "Checking Privileged Users";
            string myStatus = "Progress:";
            ProgressRecord pr = new ProgressRecord(myId, myActivity, myStatus);

            Console.WriteLine("");
            Console.WriteLine("");

            // set up domain context
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);
            //_ = new List<ExtPrincipal.UserPrincipalEx>();
            //_ = new GroupPrincipal(ctx);
            //_ = new GroupPrincipal(ctx);
            //_ = new GroupPrincipal(ctx);
            

            // Get all users from the given OU
            List<ExtPrincipal.UserPrincipalEx> AllPrivUsers = GetFromAd.GetUsersFromOU(_adminusersdn);
            WriteVerbose(string.Format("INFO - Found {0} users on {1} Organizational Unit.", AllPrivUsers.Count, _adminusersdn));
            Console.WriteLine("");

            try
            {
                // Get members from group SG_Tier0Admins
                GroupPrincipal T0Members = GroupPrincipal.FindByIdentity(ctx, "SG_Tier0Admins");

                // Get members from group SG_Tier1Admins
                GroupPrincipal T1Members = GroupPrincipal.FindByIdentity(ctx, "SG_Tier1Admins");

                // Get members from group SG_Tier2Admins
                GroupPrincipal T2Members = GroupPrincipal.FindByIdentity(ctx, "SG_Tier2Admins");

                // iterate all users
                foreach (UserPrincipal currentUser in AllPrivUsers)
                {
                    i++;

                    // Progress Record % completed
                    pr.PercentComplete = i * AllPrivUsers.Count;

                    // Process Record Current Operation
                    pr.CurrentOperation = string.Format("Procesing object # {0}: NAME: {1}", i, currentUser.DisplayName);

                    // Process Record Status message
                    pr.StatusDescription = string.Format("Processing {0} objects. Complete %: {1}", AllPrivUsers.Count, i * AllPrivUsers.Count);

                    // Write the Progress Status
                    WriteProgress(pr);

                    // make sure builtin are not processed (Admin, Guest, krbtgt...)
                    if (!(
                         (currentUser.SamAccountName.ToLower() == "thegood") ||
                         (currentUser.SamAccountName.ToLower() == "krbtgt") ||
                         (currentUser.SamAccountName.ToLower() == "theugly") ||
                         (currentUser.SamAccountName.ToLower() == "administrator") ||
                         (currentUser.SamAccountName.ToLower() == "guest") ||
                         (currentUser.SamAccountName.ToLower() == "thebad")
                        ))
                    {
                        // Get last 3 characters from SamAccountName. Should be _T0 or _T1 or _T2
                        string Last3Char = currentUser.SamAccountName.Substring(currentUser.SamAccountName.Length - 3);

                        // Select action based on last 3 char
                        switch (Last3Char.ToUpper())
                        {
                            case "_T0":
                                //Check if user is already member of group
                                if (!T0Members.Members.Contains(currentUser))
                                {
                                    // Add the user to the group
                                    T0Members.Members.Add(currentUser);
                                    T0Members.Save();
                                    WriteVerbose(string.Format("CHG - user {0} was added to {1} group", currentUser.SamAccountName, T0Members.SamAccountName));
                                    t0++;
                                }
                                break;

                            case "_T1":
                                //Check if user is already member of group
                                if (!T1Members.Members.Contains(currentUser))
                                {
                                    // Add the user to the group
                                    T1Members.Members.Add(currentUser);
                                    T1Members.Save();
                                    WriteVerbose(string.Format("CHG - user {0} was added to {1} group", currentUser.SamAccountName, T1Members.SamAccountName));
                                    t1++;
                                }
                                break;

                            case "_T2":
                                //Check if user is already member of group
                                if (!T2Members.Members.Contains(currentUser))
                                {
                                    // Add the user to the group
                                    T2Members.Members.Add(currentUser);
                                    T2Members.Save();
                                    WriteVerbose(string.Format("CHG - user {0} was added to {1} group", currentUser.SamAccountName, T2Members.SamAccountName));
                                    t2++;
                                }
                                break;

                            default:
                                WriteVerbose(string.Format("WARNING - {0} - To Be Removed from this OU.", currentUser.SamAccountName));
                                if (_disablenonstandardusers)
                                {
                                    UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(ctx, currentUser.SamAccountName.ToString());
                                    userPrincipal.Enabled = false;
                                    userPrincipal.Save();
                                    WriteVerbose(string.Format("CHG - Account {0} was disabled due compliance.", currentUser.SamAccountName));
                                }
                                break;
                        } // end switch
                    } // enf if
                } //end foreach
            } // end try
            catch
            {
            }

            pr.RecordType = ProgressRecordType.Completed;
            WriteProgress(pr);

            Console.WriteLine("");
            WriteVerbose("Added new semi- privileged users");
            WriteVerbose("--------------------------------");
            WriteVerbose(string.Format("Admin Area   / Tier0: {0}", t0));
            WriteVerbose(string.Format("Servers Area / Tier1: {0}", t1));
            WriteVerbose(string.Format("Sites Area   / Tier2: {0}", t2));
        }

        #endregion Process()

        #region End()

        /// <summary>
        ///
        /// </summary>
        protected override void EndProcessing()
        {
            if (this.MyInvocation.BoundParameters.ContainsKey("Verbose"))
            {
                string paramVerboseEnd;
                paramVerboseEnd = string.Format("\n         Function {0} finished.\n", this.MyInvocation.MyCommand);
                paramVerboseEnd += "-------------------------------------------------------------------------------\n\n";

                WriteVerbose(paramVerboseEnd);
            }
        }

        #endregion End()

        /// <summary>
        ///
        /// </summary>
        protected override void StopProcessing()
        {
        }
    } // end class
} // end namespace