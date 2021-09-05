using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Management.Automation;
using System.Security.Principal;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Clears adminCount attribute, and enables inherited security on the passed object (User or Group).</para>
    /// <para type="description">Clears adminCount attribute, and enables inherited security on the passed object (User or Group).</para>
    /// <example>
    ///     <para>This example shows how to use this CMDlet</para>
    ///     <para>-        </para>
    ///     <para>
    ///         <code>Set-AdAdminCount dvader</code>
    ///     </para>
    /// </example>
    /// <example>
    ///     <para>This example shows how to use this CMDlet using named parameters</para>
    ///     <para>-        </para>
    ///     <para>
    ///         <code>Set-AdAdminCount -SamAccountName dvader</code>
    ///     </para>
    /// </example>
    /// <remarks>Clears adminCount attribute, and enables inherited security on the passed object (User or Group).</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "AdAdminCount", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(int))]
    public class AdAdminCount : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[STRING] Sam Account Name of the object to be reseted.</para>
        ///     <para type="description">SamAccountName of the object to be reseted.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Sam Account Name of the object to be reseted."
            )]
        [ValidateNotNullOrEmpty]
        public string SamAccountName
        {
            get { return _samaccountname; }
            set { _samaccountname = value; }
        }

        private string _samaccountname;

        #endregion Parameters definition

        /// <summary>
        ///
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            //WriteVerbose(String.Format("Get all users within LDAPPath {0} which have ClearAdminCount attribute set to 1", LDAPpath));

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

                foreach(var item in pb)
                {
                    paramVerbose += string.Format("{0,-12}{1,-30}{2,-30}\n", null, item.Key, item.Value);
                }
                WriteVerbose(string.Format("{0}\n", paramVerbose));
            }

        }

        /// <summary>
        ///
        /// </summary>
        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            string result = null;

            try
            {
                result = EguibarIT.Housekeeping.Helpers.ClearAdminCount(_samaccountname);
            }
            catch { }

            WriteVerbose(result);
        }

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

        /// <summary>
        ///
        /// </summary>
        protected override void StopProcessing()
        {
        }
    }

    /// <summary>
    /// <para type="synopsis">Clears adminCount attribute, and enables inherited security on all the passed user objects.</para>
    /// <para type="description">Finds all user objects (From specific OU and childs, or All Domain), 
    /// and clears adminCount attribute, and enables inherited security 
    /// on the passed user objectlist.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters finding all users with adminCount = 1 in the domain</para>
    ///         <para>-        -</para>
    ///         <para>
    ///             <code>Set-AllUserAdminCount -all</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters finding all users with adminCount = 1 in specific OU</para>
    ///         <para>-        -</para>
    ///         <para>
    ///             <code>Set-AllUserAdminCount -SubTree -SearchRootDN "OU=Admin Accounts,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    /// </example>
    /// <remarks>Clears adminCount attribute, and enables inherited security on all the passed user objects.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "AllUserAdminCount", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(int))]
    public class AllUserAdminCount : PSCmdlet
    {

        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[SWITCH] (bool) Indicating if all users in the domain should be processed.</para>
        /// <para type="description">If present, all users in the domain should be processed</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               ParameterSetName = "AllUsers",
               HelpMessage = "Switch indicating if all users in the domain should be processed."
            )]
        [ValidateNotNullOrEmpty]
        public SwitchParameter All
        {
            get { return _all; }
            set { _all = value; }
        }
        private bool _all;

        /// <summary>
        /// <para type="inputType">[SWITCH] Indicating only a sub-tree OU should be processed.</para>
        /// <para type="description">If present, only users of a sub-tree OU should be processed.</para>
        /// </summary>
        [Parameter(
               Position = 1,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               ParameterSetName = "SubTree",
               HelpMessage = "Switch indicating only a sub-tree OU should be processed."
            )]
        [ValidateNotNullOrEmpty]
        public SwitchParameter SubTree
        {
            get { return _subtree; }
            set { _subtree = value; }
        }
        private bool _subtree;

        /// <summary>
        /// <para type="inputType">[STRING] Representing the Distinguished Name where the search starts (ej. OU=Sites,DC=EguibarIT,DC=local).</para>
        /// <para type="description">Distinguished Name where the search starts.</para>
        /// </summary>
        [Parameter(
               Position = 2,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               ParameterSetName = "SubTree",
               HelpMessage = "String representing the Distinguished Name where the search starts."
            )]
        [ValidateNotNullOrEmpty]
        public string SearchRootDN
        {
            get { return _searchrootdn; }
            set { _searchrootdn = value; }
        }
        private string _searchrootdn;

        #endregion Parameters definition

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

        /// <summary>
        ///
        /// </summary>
        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            // Define the Progress Record (Progress Bar to be displayed)
            int myId = 0;
            string myActivity = "Checking Users with adminCount attribute set to 1";
            string myStatus = "Progress:";
            ProgressRecord pr = new ProgressRecord(myId, myActivity, myStatus);

            Console.WriteLine("");
            Console.WriteLine("");

            // set up domain context
            //PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            List<ExtPrincipal.UserPrincipalEx> AllPrivUsers = new List<ExtPrincipal.UserPrincipalEx>();

            int i = 0;
            int ObjectsChanged = 0;

            try
            {
                if (_subtree)
                {
                    // Get all users from the given OU
                    //AllPrivUsers = EguibarIT.Housekeeping.GetFromAd.GetUsersFromOU(_searchrootdn);

                    // set up domain context and given OU
                    PrincipalContext ctx = new PrincipalContext(ContextType.Domain, EguibarIT.Housekeeping.AdHelper.AdDomain.GetNetbiosDomainName(), _searchrootdn);

                    //Define QueryByExample user
                    ExtPrincipal.UserPrincipalEx qbeUser = new ExtPrincipal.UserPrincipalEx(ctx)
                    {
                        // Set the value to search from
                        adminCount = "1"
                    };

                    using (PrincipalSearcher srch = new PrincipalSearcher(qbeUser))
                    {
                        foreach (ExtPrincipal.UserPrincipalEx p in srch.FindAll())
                        {
                            // Remove builtin accounts
                            if (!(
                                (p.SamAccountName.ToLower() == "administrator") ||
                                (p.SamAccountName.ToLower() == "thegood") ||
                                (p.SamAccountName.ToLower() == "theugly") ||
                                (p.SamAccountName.ToLower() == "krbtgt")
                                ))
                            {
                                AllPrivUsers.Add(p);
                            }//end if
                        }//end foreach
                    }
                }//end if
                else
                {
                    // Get all users from the domain
                    //AllPrivUsers = EguibarIT.Housekeeping.GetFromAd.GetUsersFromOU(EguibarIT.Housekeeping.AdDomain.GetrootDomainNamingContext());

                    // set up domain context
                    PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

                    //Define QueryByExample user
                    ExtPrincipal.UserPrincipalEx qbeUser = new ExtPrincipal.UserPrincipalEx(ctx)
                    {
                        // Set the value to search from
                        adminCount = "1"
                    };

                    using (PrincipalSearcher srch = new PrincipalSearcher(qbeUser))
                    {
                        foreach (ExtPrincipal.UserPrincipalEx p in srch.FindAll())
                        {
                            // Remove builtin accounts
                            if (!(
                                (p.SamAccountName.ToLower() == "administrator") ||
                                (p.SamAccountName.ToLower() == "thegood") ||
                                (p.SamAccountName.ToLower() == "theugly") ||
                                (p.SamAccountName.ToLower() == "krbtgt")
                                ))
                            {
                                AllPrivUsers.Add(p);
                            }//end if
                        }//end foreach
                    }
                }//end else

                Console.WriteLine("");
                WriteVerbose(string.Format("Iterate through each Object returned. Total objects found: {0}", AllPrivUsers.Count));
                Console.WriteLine("");

                // Iterate users
                foreach (ExtPrincipal.UserPrincipalEx p in AllPrivUsers)
                {
                    i++;

                    int PercentComplete = (i * 100 / AllPrivUsers.Count);

                    // Progress Record % completed
                    pr.PercentComplete = PercentComplete;

                    // Process Record Current Operation
                    pr.CurrentOperation = string.Format("Procesing object # {0}: NAME: {1}", i, p.DisplayName);

                    // Process Record Status message
                    pr.StatusDescription = string.Format("Processing {0} objects. Complete %: {1}", AllPrivUsers.Count, PercentComplete);

                    // Write the Progress Status
                    WriteProgress(pr);

                    // Call function to clear ClearAdminCount of current user
                    string result = EguibarIT.Housekeeping.Helpers.ClearAdminCount(p.SamAccountName);

                    if (result != null)
                    {
                        ObjectsChanged++;
                    }
                    WriteVerbose(result);
                }// End Foreach
            }//end try
            catch { }

            WriteVerbose(string.Format("{0} objects Processed. {1} objects changed", i, ObjectsChanged));
        }

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

        /// <summary>
        ///
        /// </summary>
        protected override void StopProcessing()
        {
        }
    }

    /// <summary>
    /// <para type="synopsis">Clears adminCount attribute, and enables inherited security on all the passed group objects.</para>
    /// <para type="description">Finds all group objects (From specific OU and childs, or All Domain), 
    /// and clears adminCount attribute, enabling inherited security on the 
    /// passed group objectlist.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters finding all groups with adminCount = 1 in the domain</para>
    ///         <para>
    ///             <code>Set-AllGroupAdminCount -all</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters finding all groups with adminCount = 1 in specific OU</para>
    ///         <para>
    ///             <code>Set-AllGroupAdminCount -SubTree -SearchRootDN "OU=Admin Accounts,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    /// </example>
    /// <remarks>Clears adminCount attribute, and enables inherited security on all the passed group objects.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "AllGroupAdminCount", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(int))]
    public class AllGroupAdminCount : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[SWITCH] Indicating if all groups in the domain should be processed.</para>
        /// <para type="description">If present, all groups in the domain should be processed</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               ParameterSetName = "AllGroups",
               HelpMessage = "Switch indicating if all users in the domain should be processed."
            )]
        [ValidateNotNullOrEmpty]
        public SwitchParameter All
        {
            get { return _all; }
            set { _all = value; }
        }

        private bool _all;

        /// <summary>
        /// <para type="inputType">[SWITCH] Indicating only a sub-tree OU should be processed.</para>
        /// <para type="description">If present, only groups of a sub-tree OU should be processed.</para>
        /// </summary>
        [Parameter(
               Position = 1,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               ParameterSetName = "SubTree",
               HelpMessage = "Switch indicating only a sub-tree OU should be processed."
            )]
        [ValidateNotNullOrEmpty]
        public SwitchParameter SubTree
        {
            get { return _subtree; }
            set { _subtree = value; }
        }

        private bool _subtree;

        /// <summary>
        /// <para type="inputType">[STRING] Representing the Distinguished Name where the search starts (ej. OU=Sites,DC=EguibarIT,DC=local).</para>
        /// <para type="description">Distinguished Name where the search starts.</para>
        /// </summary>
        [Parameter(
               Position = 2,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               ParameterSetName = "SubTree",
               HelpMessage = "String representing the Distinguished Name where the search starts."
            )]
        [ValidateNotNullOrEmpty]
        public string SearchRootDN
        {
            get { return _searchrootdn; }
            set { _searchrootdn = value; }
        }

        private string _searchrootdn;

        #endregion Parameters definition

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

        /// <summary>
        ///
        /// </summary>
        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            // Define the Progress Record (Progress Bar to be displayed)
            int myId = 0;
            string myActivity = "Checking Groups with adminCount attribute set to 1";
            string myStatus = "Progress:";
            ProgressRecord pr = new ProgressRecord(myId, myActivity, myStatus);

            Console.WriteLine("");
            Console.WriteLine("");

            // set up domain context
            //PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            List<ExtPrincipal.GroupPrincipalEx> AllPrivGroups = new List<ExtPrincipal.GroupPrincipalEx>();

            int i = 0;
            int ObjectsChanged = 0;

            try
            {
                if (_subtree)
                {
                    // Get all users from the given OU
                    //AllPrivUsers = EguibarIT.Housekeeping.GetFromAd.GetUsersFromOU(_searchrootdn);

                    // set up domain context and given OU
                    PrincipalContext ctx = new PrincipalContext(ContextType.Domain, EguibarIT.Housekeeping.AdHelper.AdDomain.GetNetbiosDomainName(), _searchrootdn);

                    //Define QueryByExample user
                    ExtPrincipal.GroupPrincipalEx qbeGroup = new ExtPrincipal.GroupPrincipalEx(ctx)
                    {
                        // Set the value to search from
                        adminCount = "1"
                    };


                    using (PrincipalSearcher srch = new PrincipalSearcher(qbeGroup))
                    {
                        foreach (ExtPrincipal.GroupPrincipalEx p in srch.FindAll())
                        {
                            // Remove builtin accounts (Order: Administrators, print operators, backup operators, replicator, server operators, account operators)
                            if (!(
                                (p.Sid == new SecurityIdentifier("S-1-5-32-544")) ||
                                (p.Sid == new SecurityIdentifier("S-1-5-32-550")) ||
                                (p.Sid == new SecurityIdentifier("S-1-5-32-551")) ||
                                (p.Sid == new SecurityIdentifier("S-1-5-32-552")) ||
                                (p.SamAccountName.ToLower() == "domain controllers") ||
                                (p.SamAccountName.ToLower() == "schema admins") ||
                                (p.SamAccountName.ToLower() == "enterprise admins") ||
                                (p.SamAccountName.ToLower() == "domain admins") ||
                                (p.Sid == new SecurityIdentifier("S-1-5-32-549")) ||
                                (p.Sid == new SecurityIdentifier("S-1-5-32-548")) ||
                                (p.SamAccountName.ToLower() == "read-only domain controllers")
                                ))
                            {
                                AllPrivGroups.Add(p);
                            }//end if
                        }//end foreach
                    }
                }//end if
                else
                {
                    // Get all users from the domain
                    //AllPrivUsers = EguibarIT.Housekeeping.GetFromAd.GetUsersFromOU(EguibarIT.Housekeeping.AdDomain.GetrootDomainNamingContext());

                    // set up domain context
                    PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

                    //Define QueryByExample user
                    ExtPrincipal.GroupPrincipalEx qbeGroup = new ExtPrincipal.GroupPrincipalEx(ctx)
                    {
                        // Set the value to search from
                        adminCount = "1"
                    };

                    using (PrincipalSearcher srch = new PrincipalSearcher(qbeGroup))
                    {
                        foreach (ExtPrincipal.GroupPrincipalEx p in srch.FindAll())
                        {
                            // Remove builtin accounts (Order: Administrators, print operators, backup operators, replicator, server operators, account operators)
                            if (!(
                                (p.Sid == new SecurityIdentifier("S-1-5-32-544")) ||
                                (p.Sid == new SecurityIdentifier("S-1-5-32-550")) ||
                                (p.Sid == new SecurityIdentifier("S-1-5-32-551")) ||
                                (p.Sid == new SecurityIdentifier("S-1-5-32-552")) ||
                                (p.SamAccountName.ToLower() == "domain controllers") ||
                                (p.SamAccountName.ToLower() == "schema admins") ||
                                (p.SamAccountName.ToLower() == "enterprise admins") ||
                                (p.SamAccountName.ToLower() == "domain admins") ||
                                (p.Sid == new SecurityIdentifier("S-1-5-32-549")) ||
                                (p.Sid == new SecurityIdentifier("S-1-5-32-548")) ||
                                (p.SamAccountName.ToLower() == "read-only domain controllers")
                                ))
                            {
                                AllPrivGroups.Add(p);
                            }//end if
                        }//end foreach
                    }
                }//end else

                Console.WriteLine("");
                WriteVerbose(string.Format("Iterate through each Object returned. Total objects found: {0}", AllPrivGroups.Count));
                Console.WriteLine("");

                // Iterate users
                foreach (ExtPrincipal.GroupPrincipalEx p in AllPrivGroups)
                {
                    i++;

                    int PercentComplete = (i * 100 / AllPrivGroups.Count);

                    // Progress Record % completed
                    pr.PercentComplete = PercentComplete;

                    // Process Record Current Operation
                    pr.CurrentOperation = string.Format("Procesing object # {0}: NAME: {1}", i, p.DisplayName);

                    // Process Record Status message
                    pr.StatusDescription = string.Format("Processing {0} objects. Complete %: {1}", AllPrivGroups.Count, PercentComplete);

                    // Write the Progress Status
                    WriteProgress(pr);

                    // Call function to clear ClearAdminCount of current user
                    string result = EguibarIT.Housekeeping.Helpers.ClearAdminCount(p.SamAccountName);

                    if (result != null)
                    {
                        ObjectsChanged++;
                    }
                    result = null;

                    WriteVerbose(result);
                }// End Foreach
            }//end try
            catch { }

            WriteVerbose(string.Format("{0} objects Processed. {1} objects changed", i, ObjectsChanged));
        }

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

        /// <summary>
        ///
        /// </summary>
        protected override void StopProcessing()
        {
        }
    }
}