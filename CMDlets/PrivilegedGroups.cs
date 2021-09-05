using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Function for Privileged Groups housekeeping</para>
    /// <para type="description">Ensure that any group stored on Admin OU (Tier0) contains only authorized users. 
    /// An Authorized user is a user created and mantained by the Tier0 Admins, and is 
    /// usually identified by the SamAccountName suffix of either _T0, or _T1 or _T2. 
    /// Any non authorized user will be inmediatelly revoked from these groups.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet.</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-PrivilegedGroupsHousekeeping "OU=Groups,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///    </example>
    ///    <example>
    ///         <para>This example shows how to use this CMDlet using named parameters.</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-PrivilegedGroupsHousekeeping -AdminUsersDN "OU=Groups,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///    </example>
    ///    <example>
    ///         <para>This example shows how to use this CMDlet using named parameters and provides a exclusion list.</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-PrivilegedGroupsHousekeeping -AdminUsersDN "OU=Groups,OU=Admin,DC=EguibarIT,DC=local" -ExcludeList "dvader", "hsolo"</code>
    ///         </para>
    ///    </example>
    /// <remarks>Ensures only Privileged and Semi-Privileged users are members of Admin groups.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "PrivilegedGroupsHousekeeping", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(int))]
    public class PrivilegedGroups : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[STRING] Admin Users OU Distinguished Name. (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        /// <para type="description">Distinguished Name of the container where the users are located.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Admin Users OU Distinguished Name."
            )]
        [ValidateNotNullOrEmpty]
        public string AdminUsersDN
        {
            get { return _adminusersdn; }
            set { _adminusersdn = value; }
        }

        private string _adminusersdn;

        /// <summary>
        /// <para type="inputType">List[STRING] user list array</para>
        /// <para type="description">Userlist to be excluded from this process.</para>
        /// </summary>
        [Parameter(
               Position = 1,
               Mandatory = false,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Userlist to be excluded from this process."
            )]
        public List<string> ExcludeList
        {
            get { return _excludelist; }
            set { _excludelist = value; }
        }

        private List<string> _excludelist;

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

            // Item Counter
            int i = 0;

            // Removed Users counter
            int UserRemoved = 0;

            // Declare the list of groups
            List<GroupPrincipal> groups = new List<GroupPrincipal>();
            // Declare final list of exclusions
            List<string> _NewExcludeList = new List<string>();

            // Define the Progress Record (Progress Bar to be displayed)
            int myId = 0;
            string myActivity = "Checking Privileged/Semi-Privileged Groups";
            string myStatus = "Progress:";
            ProgressRecord pr = new ProgressRecord(myId, myActivity, myStatus);

            try
            {
                // set up domain context and given OU
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, EguibarIT.Housekeeping.AdHelper.AdDomain.GetNetbiosDomainName(), _adminusersdn);

                GroupPrincipal qbeGroup = new GroupPrincipal(ctx);

                using (PrincipalSearcher srch = new PrincipalSearcher(qbeGroup))
                {
                    // iterate the results
                    foreach (GroupPrincipal p in srch.FindAll())
                    {
                        // Add the current group to the list.
                        groups.Add(p);
                    }
                }//end using

                // Get  Other PRIVILEGED groups

                // set up domain context
                ctx = new PrincipalContext(ContextType.Domain);

                // find the ADMINISTRATORS group and add it to the group list.
                groups.Add(GroupPrincipal.FindByIdentity(ctx, "Administrators"));

                // find the Account Operators group and add it to the group list.
                groups.Add(GroupPrincipal.FindByIdentity(ctx, "Account Operators"));

                // find the Backup Operators group and add it to the group list.
                groups.Add(GroupPrincipal.FindByIdentity(ctx, "Backup Operators"));

                // find the Hyper-V Administrators group and add it to the group list.
                groups.Add(GroupPrincipal.FindByIdentity(ctx, "Hyper-V Administrators"));

                // find the Incoming Forest Trust Builders group and add it to the group list.
                groups.Add(GroupPrincipal.FindByIdentity(ctx, "Incoming Forest Trust Builders"));

                // find the Server Operators group and add it to the group list.
                groups.Add(GroupPrincipal.FindByIdentity(ctx, "Server Operators"));
            }// end try
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString();
                E.Message.ToString();
                //WriteObject("ERROR - Something went wrong while adding MSA to the group.");
            }

            // Total Objects Found
            // Total Objects Found
            int TotalObjectsFound = groups.Count;

            // Check if exclusion list is provided
            if (_excludelist != null)
            {
                // If the Administrator does not exist, add it
                if (!_excludelist.Contains("Administrator"))
                {
                    _excludelist.Add("Administrator");
                }

                // If the TheGood does not exist, add it
                if (!_excludelist.Contains("TheGood"))
                {
                    _excludelist.Add("TheGood");
                }

                // If the TheUgly does not exist, add it
                if (!_excludelist.Contains("TheUgly"))
                {
                    _excludelist.Add("TheUgly");
                }

                // If the krbtgt does not exist, add it
                if (!_excludelist.Contains("krbtgt"))
                {
                    _excludelist.Add("krbtgt");
                }

                // If the DefaultAccount does not exist, add it
                if (!_excludelist.Contains("DefaultAccount"))
                {
                    _excludelist.Add("DefaultAccount");
                }

                // Copy list to new list
                _NewExcludeList.AddRange(_excludelist);
            } //en if
            else
            {
                _NewExcludeList.Add("Administrator");
                _NewExcludeList.Add("TheGood");
                _NewExcludeList.Add("TheUgly");
                _NewExcludeList.Add("krbtgt");
                _NewExcludeList.Add("DefaultAccount");
            }

            Console.WriteLine("");
            Console.WriteLine("");
            WriteVerbose(string.Format("Iterate through each group returned. Total groups found: {0}", TotalObjectsFound));
            Console.WriteLine("");

            try
            {
                // Iterate through all groups returned
                foreach (GroupPrincipal group in groups)
                {
                    i++;

                    int PercentComplete = (i * 100 / TotalObjectsFound);

                    // Progress Record % completed
                    pr.PercentComplete = PercentComplete;

                    // Process Record Current Operation
                    pr.CurrentOperation = string.Format("Procesing object # {0}: NAME: {1}", i, group.DisplayName);

                    // Process Record Status message
                    pr.StatusDescription = string.Format("Processing {0} objects. Complete %: {1}", TotalObjectsFound, PercentComplete);

                    // Write the Progress Status
                    WriteProgress(pr);

                    //Exclude ServiceAccount Groups & Domain Users; Domain Computers; Domain Guests
                    if (!(
                        (group.SamAccountName == "SG_T0SA") ||
                        (group.SamAccountName == "SG_T1SA") ||
                        (group.SamAccountName == "SG_T2SA") ||
                        (group.SamAccountName == "Domain Users") ||
                        (group.SamAccountName == "Domain Computers") ||
                        (group.SamAccountName == "Domain Guests")
                        ))
                    {
                        // Get the group membership
                        List<Principal> users = EguibarIT.Housekeeping.GetFromAd.GetGroupMembers(group.SamAccountName.ToString());

                        // Iterate group membership
                        foreach (Principal currentUser in users)
                        {
                            // Select only users
                            if (currentUser is UserPrincipal)
                            {
                                // Check the exclude list
                                if (!_NewExcludeList.Contains(currentUser.SamAccountName.ToString()))
                                {
                                    // Get last 3 characters from SamAccountName. Should be _T0 or _T1 or _T2
                                    string Last3Char = currentUser.SamAccountName.Substring(currentUser.SamAccountName.Length - 3);

                                    // Select action based on last 3 char
                                    if (!(
                                        (Last3Char == "_T0") ||
                                        (Last3Char == "_T1") ||
                                        (Last3Char == "_T2")
                                        ))
                                    {
                                        try
                                        {
                                            // Remove current user from group
                                            group.Members.Remove(currentUser);
                                            group.Save();

                                            // Provide verbose message
                                            WriteWarning(string.Format("CHG - Account \"{0}\" was removed from \"{1}\" group. This is because the account is NOT a Privileged or Semi-Privileged user. Only _T0, _T1 or _T2 accounts are permited to be member of these groups.", currentUser.SamAccountName, group.SamAccountName));
                                            Console.WriteLine("");

                                            // Increase user removed counter
                                            UserRemoved++;
                                        } // end try
                                        catch (System.DirectoryServices.DirectoryServicesCOMException E)
                                        {
                                            //doSomething with E.Message.ToString();
                                            E.Message.ToString();
                                            WriteObject("ERROR - Something went wrong while removing object from the group.");
                                        }
                                    } // end if
                                }// end if
                            } // end if
                        } // end foreach
                    } //end if
                }
            } //end try
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                E.Message.ToString();
            }

            pr.RecordType = ProgressRecordType.Completed;
            WriteProgress(pr);

            Console.WriteLine("");
            WriteVerbose("A Semi-Privileged and/or Privileged group can ONLY contain standard accounts (_T0 or _T1 or _T2.");
            WriteVerbose("Any userID which does not complies with this statement, will automatically be removed from the group");
            WriteVerbose("---------------------------------------------------------------------------------------------------");
            WriteVerbose(string.Format("{0} users were removed from Privileged or Semi-Privileged groups.", UserRemoved));
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