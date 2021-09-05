using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Function for ServiceAccount housekeeping</para>
    /// <para type="description">For each of the service accounts stored within the SA container, make it part  
    /// of the corresponding service account group based on its tier. 
    /// For example, the container of Tier0 service accounts must belong to the group SG_T0SA.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ServiceAccountHousekeeping "OU=T0,OU=SA,OU=Admin,DC=EguibarIT,DC=local" SG_T0SA</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ServiceAccountHousekeeping -ServiceAccountDN "OU=T0,OU=SA,OU=Admin,DC=EguibarIT,DC=local" -SAGroupName SG_T0SA</code>
    ///         </para>
    /// </example>
    /// <remarks>Ensure each Service Account is member of the corresponding tier group.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "ServiceAccountHousekeeping", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(int))]
    public class ServiceAccountHousekeeping : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[STRING] representing the Distinguished Name of the object (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        /// <para type="description">Distinguished Name of the container where the Service Accounts are located.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Distinguished Name of the container where the Service Accounts are located."
            )]
        [ValidateNotNullOrEmpty]
        public string ServiceAccountDN
        {
            get { return _serviceaccountdn; }
            set { _serviceaccountdn = value; }
        }

        private string _serviceaccountdn;

        /// <summary>
        /// <para type="inputType">[STRING] Name of the corresponding tier service account group (For tier0: SG_T0SA; for Tier1: SG_T1SA; for Tier2: SG_T2SA)</para>
        /// <para type="description">Name of the corresponding tier service account group (For tier0: SG_T0SA; for Tier1: SG_T1SA; for Tier2: SG_T2SA).</para>
        /// </summary>
        [Parameter(
               Position = 1,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Name of the corresponding tier service account group (For tier0: SG_T0SA; for Tier1: SG_T1SA; for Tier2: SG_T2SA)."
            )]
        [ValidateNotNullOrEmpty]
        [ValidateSet("SG_T0SA", "SG_T1SA", "SG_T2SA", IgnoreCase = true)]
        public string SAGroupName
        {
            get { return _sagroupname; }
            set { _sagroupname = value; }
        }

        private string _sagroupname;

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

            try
            {
                // set up domain context
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

                List<ExtPrincipal.UserPrincipalEx> users = new List<ExtPrincipal.UserPrincipalEx>();
                List<Principal> MSAs = new List<Principal>();

                // Get the AD group
                GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, _sagroupname);

                Console.WriteLine("");
                Console.WriteLine("");

                // Get users from the given OU
                users = EguibarIT.Housekeeping.GetFromAd.GetUsersFromOU(_serviceaccountdn);
                WriteVerbose(string.Format("INFO - Found {0} users on {1} Organizational Unit.", users.Count, _serviceaccountdn));

                MSAs = EguibarIT.Housekeeping.GetFromAd.GetMSAsFromOU(_serviceaccountdn);
                WriteVerbose(string.Format("INFO - Found {0} MSA on {1} Organizational Unit.", MSAs.Count, _serviceaccountdn));

                Console.WriteLine("");

                try
                {
                    // Iterate members to check if userObject already member of the group
                    foreach (UserPrincipal currentUser in users)
                    {
                        if (currentUser != null)
                        {
                            // check if user is member of that group
                            if (currentUser.IsMemberOf(group))
                            {
                                WriteVerbose(string.Format("INFO - {0} is already member of {1} group", currentUser.SamAccountName, group.SamAccountName));
                            }
                            else
                            {
                                // Add the missing member
                                group.Members.Add(ctx, IdentityType.SamAccountName, currentUser.SamAccountName);
                                group.Save();
                                WriteVerbose(string.Format("CHG - user {0} was added to {1} group", currentUser.SamAccountName, group.SamAccountName));
                            }
                        }
                    }
                }// end try
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    //doSomething with E.Message.ToString();
                    E.Message.ToString();
                    WriteObject("ERROR - Something went wrong while adding MSA to the group.");
                }

                try
                {
                    // Iterate members to check if MSA already member of the group
                    foreach (Principal currentMSA in MSAs)
                    {
                        if (currentMSA != null)
                        {
                            // check if user is member of that group
                            if (currentMSA.IsMemberOf(group))
                            {
                                WriteVerbose(string.Format("INFO - {0} is already member of {1} group", currentMSA.SamAccountName, group.SamAccountName));
                            }
                            else
                            {
                                // Add the missing member
                                group.Members.Add(ctx, IdentityType.SamAccountName, currentMSA.SamAccountName);
                                group.Save();
                                WriteVerbose(string.Format("CHG - MSA {0} was added to {1} group", currentMSA.SamAccountName, group.SamAccountName));
                            }
                        }
                    }
                }// end try
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    //doSomething with E.Message.ToString();
                    E.Message.ToString();
                    WriteObject("ERROR - Something went wrong while adding MSA to the group.");
                }

                try
                {
                    // Get group members
                    List<Principal> groupMembers = EguibarIT.Housekeeping.GetFromAd.GetGroupMembers(_sagroupname);

                    // Iterate accounts of the given OU (Tier OU) and compare it to group membership. Remove in case is member and not on OU
                    foreach (ExtPrincipal.UserPrincipalEx currentOUusers in groupMembers)
                    {
                        if (currentOUusers != null)
                        {
                            if (!(users.Contains(currentOUusers) || MSAs.Contains(currentOUusers)))
                            {
                                // Add the missing member Set-ServiceAccountHousekeeping -LDAPpath 'OU=T0SA,OU=Service Accounts,OU=Admin,DC=EguibarIT,DC=local' -SAGroupName SG_T0SA -Verbose
                                group.Members.Remove(ctx, IdentityType.SamAccountName, currentOUusers.SamAccountName);
                                group.Save();
                                WriteVerbose(string.Format("CHG - Account {0} was removed from {1} group. This is because the account is not present anymore on the OU. Only accounts on the corresponding OU (Tier0, Tier1 or Tier2) will be member of the corresponding group (SG_T0SA, SG_T1SA or SG_T2SA", currentOUusers.SamAccountName, group.SamAccountName));
                            }
                        }
                    }
                }// end try
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    //doSomething with E.Message.ToString();
                    E.Message.ToString();
                    WriteObject("ERROR - Something went wrong while removing object from the group.");
                }
            } // end try
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString();
                E.Message.ToString();
                WriteObject("ERROR - Something went wrong when executing the function.");
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
    } // end class
} // end namespace