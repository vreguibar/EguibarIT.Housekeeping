using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Integrity check of Semi-Privileged accounts.</para>
    /// <para type="description">A Semi-Privileged account can only exist if a Non-Privileged account (standard user) does exists. This is called Key-Pair. 
    /// If the key-pair non-privileged account (standard user) gets disabled, the semi-privileged account must be disabled as well. 
    /// If the key-pair non-privileged account (standard user) does not exist, the Semi-Privileged account must be deleted inmediatelly. 
    /// This function will read all Semi-Privileged accounts and will search its key-pair (SID of the non privileged account), 
    /// applying the mentioned actions. 
    /// For this check to work properly, all semi-privileged accounts should have the non-privileged SID stored on the 
    /// employeeNumber attribute.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet. </para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-SemiPrivilegedKeyPairCheck "OU=Users,OU=Admin,DC=EguibarIT,DC=local" "TheGood", "TheUgly"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters. </para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-SemiPrivilegedKeyPairCheck -AdminUsersDN "OU=Users,OU=Admin,DC=EguibarIT,DC=local" -ExcludeList "TheGood", "TheUgly"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using splatting. </para>
    ///         <para>-        </para>
    ///         <para>
    ///         <code>$params = @{
    ///             AdminUsersDN = "OU=Users,OU=Admin,DC=EguibarIT,DC=local"
    ///             ExcludeList  = "TheGood", "TheUgly"
    ///             Verbose      = $True
    ///         }
    ///         Set-SemiPrivilegedKeyPairCheck @params
    ///         </code>
    ///         </para>
    ///     </example>
    /// <remarks>Ensures only each Semi-Privileged users does have a valid Non-Privileged user.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "SemiPrivilegedKeyPairCheck", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(string))]
    public class SemiPrivilegedKeyPairCheck : PSCmdlet
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

            // Declare final list of exclusions
            List<string> _NewExcludeList = new List<string>();

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

                // If the DefaultAccount does not exist, add it
                if (!_excludelist.Contains("TheBad"))
                {
                    _excludelist.Add("TheBad");
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
                _NewExcludeList.Add("TheBad");
            }

            // Define the Progress Record (Progress Bar to be displayed)
            int myId = 0;
            string myActivity = "Checking Semi-Privileged Users";
            string myStatus = "Progress:";
            ProgressRecord pr = new ProgressRecord(myId, myActivity, myStatus);

            Console.WriteLine("");
            Console.WriteLine("");

            // set up domain context
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            // Disable List
            List<string> DisableList = new List<string>();

            // Delete List
            List<string> DeleteList = new List<string>();

            // Flags to disable and/or delete
            bool DisableSemiPrivilegedUser = false;
            bool DeleteSemiPrivilegedUser = false;

            List<ExtPrincipal.UserPrincipalEx> AllPrivUsers = new List<ExtPrincipal.UserPrincipalEx>();

            // Get all users from the given OU
            AllPrivUsers = EguibarIT.Housekeeping.GetFromAd.GetUsersFromOU(_adminusersdn);

            Console.WriteLine("");
            if (AllPrivUsers.Count > 0)
            {
                WriteVerbose(string.Format("INFO - Found {0} Privileged/Semi-Privileged users.", AllPrivUsers.Count));
            }
            else
            {
                WriteVerbose("INFO - No Privileged/Semi-Privileged users found on the given Distinguished Name container.");
            }

            //Iterate through all objects found
            foreach (ExtPrincipal.UserPrincipalEx semiPrivilegedUser in AllPrivUsers)
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

                //Non-Privileged user
                ExtPrincipal.UserPrincipalEx NPuser = new ExtPrincipal.UserPrincipalEx(ctx);

                //Exclude list from the process
                if (!_NewExcludeList.Contains(semiPrivilegedUser.SamAccountName))
                {

                    //Check if attribute employeeNumber contains data
                    if (semiPrivilegedUser.employeeNumber == null)
                    {
                        // Asume this is not a valid Semi-Privileged user
                        DeleteSemiPrivilegedUser = true;

                        WriteWarning(string.Format("WARNING - User {0} not linked to non-privileged user. Attribute empty! Account will be deleted.", semiPrivilegedUser.SamAccountName));
                    }//end if
                    else
                    {
                        //Get the NonPrivileged user

                        NPuser = ExtPrincipal.UserPrincipalEx.FindByIdentity(ctx, IdentityType.Sid, semiPrivilegedUser.employeeNumber);

                        if(NPuser == null)
                        {
                            WriteWarning(string.Format("WARNING - User {0} not linked to a non-privileged user. Standard user NOT FOUND! Account will be deleted.", semiPrivilegedUser.SamAccountName));

                            // Exception indicating user does not exists
                            DeleteSemiPrivilegedUser = true;
                        }
                        else
                        {
                            WriteVerbose(string.Format("INFO - Non-privileged user found: {0}.", NPuser.SamAccountName));

                            // Standard used DOES exist. Do not delete.
                            DeleteSemiPrivilegedUser = false;
                        }

                    } //end if

                    if (!(NPuser == null))
                    {
                        //Check if NonPrivileged user is disabled
                        if (NPuser.Enabled != true)
                        {
                            DisableSemiPrivilegedUser = true;
                            WriteVerbose(string.Format("INFO - Non-Privileged user {0} is disabled.", NPuser.SamAccountName));
                        }
                    }

                    // Disable SemiPrivileged Account
                    if (DisableSemiPrivilegedUser)
                    {
                        Console.WriteLine("");
                        WriteVerbose(string.Format("Non-Privileged user {0} is disabled. The Semi-Privileged account {1} will be disabled.", NPuser.SamAccountName, semiPrivilegedUser.SamAccountName));
                        try
                        {
                            semiPrivilegedUser.Enabled = false;
                            semiPrivilegedUser.Save();

                            DisableList.Add(semiPrivilegedUser.SamAccountName.ToString());
                        }//end try
                        catch (Exception Ex)
                        {
                            WriteWarning(string.Format("ERROR - Something went wrong when trying to disable {0} user", semiPrivilegedUser.SamAccountName));

                            Console.WriteLine(Ex.Message);
                        }
                    }//end if

                    // Delete SemiPrivileged Account
                    if (DeleteSemiPrivilegedUser)
                    {
                        Console.WriteLine("");
                        WriteVerbose(string.Format("Non-Privileged user does not exist. The Semi-Privileged account {0} will be Deleted.", semiPrivilegedUser.SamAccountName));
                        try
                        {
                            DeleteList.Add(semiPrivilegedUser.SamAccountName.ToString());

                            semiPrivilegedUser.Delete();
                        }//end try
                        catch (Exception Ex)
                        {
                            WriteWarning(string.Format("ERROR - Something went wrong when trying to delete {0} user", semiPrivilegedUser.SamAccountName));

                            Console.WriteLine(Ex.Message);
                        }
                    }//end if
                }//end if

                DisableSemiPrivilegedUser = false;
                DeleteSemiPrivilegedUser = false;

                if (!(NPuser == null))
                {
                    NPuser.Dispose();
                }

            }//end foreach

            Console.WriteLine("");
            Console.WriteLine("");
            WriteVerbose("The following Semi-Privileged users where disabled or deleted");
            Console.WriteLine("");
            WriteVerbose(string.Format("Total disabled users...: {0}", DisableList.Count));
            DisableList.ForEach(WriteVerbose);
            Console.WriteLine("");
            WriteVerbose(string.Format("Total deleted users...: {0}", DeleteList.Count));
            DeleteList.ForEach(WriteVerbose);


            ctx.Dispose();
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
}//end namespace