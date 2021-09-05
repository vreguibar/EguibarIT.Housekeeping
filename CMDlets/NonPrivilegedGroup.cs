using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Function for consistency check on semi-privileged accounts</para>
    /// <para type="description">For each of the user accounts stored within the USERS container of Admin Area, all groups will be enumerated 
    /// and evaluated.In case that any of the groups the user belongs to are not part of Admin Area and/or BuiltIn container, 
    /// will be removed from the user.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-NonPrivilegedGroupHousekeeping "OU=Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-NonPrivilegedGroupHousekeeping -AdminUsersDN "OU=Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    /// </example>
    /// <remarks>Ensures only Privileged and Semi-Privileged users are used within Admin groups.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "NonPrivilegedGroupHousekeeping", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(int))]
    public class NonPrivilegedGroup : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[STRING] Admin User Account OU Distinguished Name. (ej. OU=Users,OU=Admin,DC=EguibarIT,DC=local).</para>
        /// <para type="description">Distinguished Name of the container where the users are located.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Admin User Account OU Distinguished Name."
            )]
        [ValidateNotNullOrEmpty]
        public string AdminUsersDN
        {
            get { return _adminusersdn; }
            set { _adminusersdn = value; }
        }

        private string _adminusersdn;

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

            int i = 0;

            // Define the Progress Record (Progress Bar to be displayed)
            int myId = 0;
            string myActivity = "Checking Non-Privileged Groups";
            string myStatus = "Progress:";
            ProgressRecord pr = new ProgressRecord(myId, myActivity, myStatus);

            Console.WriteLine("");
            Console.WriteLine("");
            
            // Get all users from the given OU
            List<ExtPrincipal.UserPrincipalEx> AllPrivUsers = EguibarIT.Housekeeping.GetFromAd.GetUsersFromOU(_adminusersdn);
            WriteVerbose(string.Format("INFO - Found {0} Privileged/Semi-Privileged users.", AllPrivUsers.Count));
            Console.WriteLine("");

            try
            {
                //Iterate through all semi-privileged users
                foreach (Principal p in AllPrivUsers)
                {
                    i++;

                    int PercentComplete = (i * 100 / AllPrivUsers.Count);

                    // Progress Record % completed
                    pr.PercentComplete = PercentComplete;

                    // Process Record Current Operation
                    pr.CurrentOperation = string.Format("Procesing object # {0}: NAME: {1}", i, AllPrivUsers.Count);

                    // Process Record Status message
                    pr.StatusDescription = string.Format("Processing {0} objects. Complete %: {1}", AllPrivUsers.Count, PercentComplete);

                    // Write the Progress Status
                    WriteProgress(pr);

                    //Iterate through the list of groups of the current user
                    foreach (GroupPrincipal gp in p.GetGroups())
                    {
                        //Check the distinguished name of the group. If not part of Admin area and/or BuiltIn continue
                        if (!(
                            gp.DistinguishedName.Contains("OU=Admin") ||
                            gp.DistinguishedName.Contains("CN=Builtin") ||
                            gp.SamAccountName.Contains("Domain Users")
                            ))
                        {
                            //Remove the user from the non-privileged group.
                            gp.Members.Remove(p);
                            gp.Save();

                            WriteVerbose(string.Format("CHG - Account {0} was removed from {1} group. Privileged or Semi-Privileged accounts cannot be members of a Non-Privileged group. Privileged or Semi-Privileged accounts can only be members of Privileged/Semi-Privileged groups", p.SamAccountName, gp.SamAccountName));
                            Console.WriteLine("");
                        }//end if
                    }//end foreach
                }//end foreach
            } //end try
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                E.Message.ToString();
            }
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
    }//end class
}