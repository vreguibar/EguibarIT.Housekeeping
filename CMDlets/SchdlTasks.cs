using System;
using System.Collections.Generic;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Create a new Scheduled Tasks which executes Set-AllUserAdminCount</para>
    /// <para type="description">Create a new Scheduled Tasks, which will trigger 12 times per day. 
    /// The Action configured is a PowerShell function from the Housekeeping module: Set-AllUserAdminCount.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleAllUserAdminCount -ServiceAccount "SA_R2D2"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleAllUserAdminCount "SA_R2D2"</code>
    ///         </para>
    ///     </example>
    /// <remarks>Create new Scheduled Task which resets All Users Admin Count.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "ScheduleAllUserAdminCount", ConfirmImpact = ConfirmImpact.Low)]
    [OutputType(typeof(int))]
    public class ScheduleAllUserAdminCount : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[String] SamAccountName of the group managed service account used to execute the scheduled task.</para>
        ///     <para type="description">SamAccountName of the service account (or gMSA) used to execute the scheduled task.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[String] SamAccountName of the group managed service account used to execute the scheduled task."
            )]
        [ValidateNotNullOrEmpty]
        public string ServiceAccount
        {
            get { return _serviceaccount; }
            set { _serviceaccount = value; }
        }

        private string _serviceaccount;

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

            // Set variables
            string TaskName = "All Users Admin Count";
            string Description = "Finds all user objects in the current Domain, clears the adminCount attribute, and enables inherited security.";

            string PsArg1 = "-Command \" & { Import-Module EguibarIT.Housekeeping; ";
            string PsArg2 = "Set-AllUserAdminCount -all; ";
            string PsArg3 = "exit $LASTEXITCODE }\"";
            string PsArguments = PsArg1 + PsArg2 + PsArg3;

            string ExeActionID = "opens powershell with parameters";
            EguibarIT.Housekeeping.AdHelper.Ocurrences Ocurrences = EguibarIT.Housekeeping.AdHelper.Ocurrences.Twelve;

            // Call the function to create task
            EguibarIT.Housekeeping.TaskScheduler.NewScheduledTask(TaskName, _serviceaccount, Description, PsArguments, ExeActionID, Ocurrences);
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

    /// <summary>
    /// <para type="synopsis">Create a new Scheduled Tasks which executes Set-AllGroupAdminCount</para>
    /// <para type="description">Create a new Scheduled Tasks, which will trigger 12 times per day. 
    /// The Action configured is a PowerShell function from the Housekeeping module: Set-AllGroupAdminCount.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleAllGroupAdminCount -ServiceAccount "SA_R2D2"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///         <code>Set-ScheduleAllGroupAdminCount "SA_R2D2"</code>
    ///         </para>
    ///     </example>
    /// <remarks>Create new Scheduled Task which resets All Groups Admin Count.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "ScheduleAllGroupAdminCount", ConfirmImpact = ConfirmImpact.Low)]
    [OutputType(typeof(int))]
    public class ScheduleAllGroupAdminCount : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[String] SamAccountName of the group managed service account used to execute the scheduled task.</para>
        ///     <para type="description">SamAccountName of the service account (or gMSA) used to execute the scheduled task.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[String] SamAccountName of the group managed service account used to execute the scheduled task."
            )]
        [ValidateNotNullOrEmpty]
        public string ServiceAccount
        {
            get { return _serviceaccount; }
            set { _serviceaccount = value; }
        }

        private string _serviceaccount;

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

            // Set variables
            string TaskName = "All Groups Admin Count";
            string Description = "Finds all group objects in the current Domain, clears the adminCount attribute, and enables inherited security.";

            string PsArg1 = "-Command \" & { Import-Module EguibarIT.Housekeeping; ";
            string PsArg2 = "Set-AllGroupAdminCount -all; ";
            string PsArg3 = "exit $LASTEXITCODE }\"";
            string PsArguments = PsArg1 + PsArg2 + PsArg3;

            string ExeActionID = "opens powershell with parameters";
            EguibarIT.Housekeeping.AdHelper.Ocurrences Ocurrences = EguibarIT.Housekeeping.AdHelper.Ocurrences.Twelve;

            // Call the function to create task
            EguibarIT.Housekeeping.TaskScheduler.NewScheduledTask(TaskName, _serviceaccount, Description, PsArguments, ExeActionID, Ocurrences);
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

    /// <summary>
    /// <para type="synopsis">Create a new Scheduled Tasks which executes Set-PrivilegedUsersHousekeeping</para>
    /// <para type="description">Create a new Scheduled Tasks, which will trigger 12 times per day. 
    /// The Action configured is a PowerShell function from the Housekeeping module: Set-PrivilegedUsersHousekeeping</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-SchedulePrivilegedUsers -ServiceAccount "SA_R2D2" -AdminUsersDN "OU=Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-SchedulePrivilegedUsers "SA_R2D2" "OU=Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    /// <remarks>Create new Scheduled Task which ensure each Semi-Privileged account is member of the corresponding tier group.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "SchedulePrivilegedUsers", ConfirmImpact = ConfirmImpact.Low)]
    [OutputType(typeof(int))]
    public class SchedulePrivilegedUsers : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[String] SamAccountName of the group managed service account used to execute the scheduled task.</para>
        ///     <para type="description">SamAccountName of the service account (or gMSA) used to execute the scheduled task.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[String] SamAccountName of the group managed service account used to execute the scheduled task."
            )]
        [ValidateNotNullOrEmpty]
        public string ServiceAccount
        {
            get { return _serviceaccount; }
            set { _serviceaccount = value; }
        }

        private string _serviceaccount;

        /// <summary>
        ///     <para type="inputType">[STRING] Admin User Account OU Distinguished Name. (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        ///     <para type="description">Distinguished Name of the container where the Admin users are located.</para>
        /// </summary>
        [Parameter(
               Position = 1,
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

            // Set variables
            string TaskName = "Privileged Users";
            string Description = "Privileged Users housekeeping. For each of the user accounts stored within the USERS container, make it part of the corresponding admin tier group. For example, if the user account is a Tier2 account(_T2), the user must belong to SG_Tier2Admins group.";

            string PsArg1 = "-Command \" & { Import-Module EguibarIT.Housekeeping; ";
            string PsArg2 = string.Format("Set-PrivilegedUsersHousekeeping -AdminUsersDN \"{0}\" -DisableNonStandardUsers; ", _adminusersdn);
            string PsArg3 = "exit $LASTEXITCODE }\"";
            string PsArguments = PsArg1 + PsArg2 + PsArg3;

            string ExeActionID = "opens powershell with parameters";
            EguibarIT.Housekeeping.AdHelper.Ocurrences Ocurrences = EguibarIT.Housekeeping.AdHelper.Ocurrences.Twelve;

            // Call the function to create task
            EguibarIT.Housekeeping.TaskScheduler.NewScheduledTask(TaskName, _serviceaccount, Description, PsArguments, ExeActionID, Ocurrences);
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

    /// <summary>
    /// <para type="synopsis">Create a new Scheduled Tasks which executes Set-NonPrivilegedGroupHousekeeping</para>
    /// <para type="description">Create a new Scheduled Tasks, which will trigger 12 times per day. The Action configured 
    /// is a PowerShell function from the Housekeeping module: Set-NonPrivilegedGroupHousekeeping</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleNonPrivilegedGroups -ServiceAccount "SA_R2D2" -AdminUsersDN "OU=Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleNonPrivilegedGroups "SA_R2D2" "OU=Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    /// <remarks>Create new Scheduled Task which ensures only Privileged and Semi-Privileged users are used within Admin groups.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "ScheduleNonPrivilegedGroups", ConfirmImpact = ConfirmImpact.Low)]
    [OutputType(typeof(int))]
    public class ScheduleNonPrivilegedGroups : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[String] SamAccountName of the group managed service account used to execute the scheduled task.</para>
        ///     <para type="description">SamAccountName of the service account (or gMSA) used to execute the scheduled task.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[String] SamAccountName of the group managed service account used to execute the scheduled task."
            )]
        [ValidateNotNullOrEmpty]
        public string ServiceAccount
        {
            get { return _serviceaccount; }
            set { _serviceaccount = value; }
        }

        private string _serviceaccount;

        /// <summary>
        ///     <para type="inputType">[STRING] Admin User Account OU Distinguished Name. (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        ///     <para type="description">Distinguished Name of the container where the Admin users are located.</para>
        /// </summary>
        [Parameter(
               Position = 1,
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

            // Set variables
            string TaskName = "Non-Privileged Groups";
            string Description = "Consistency check on semi-privileged accounts. For each of the user accounts stored within the USERS container of Admin Area, all groups will be enumerated and evaluated. In case that any of the groups the user belongs to are not part of Admin Area and/ or BuiltIn container, will be removed from the user.";

            string PsArg1 = "-Command \" & { Import-Module EguibarIT.Housekeeping; ";
            string PsArg2 = string.Format("Set-NonPrivilegedGroupHousekeeping -AdminUsersDN \"{0}\"; ", _adminusersdn);
            string PsArg3 = "exit $LASTEXITCODE }\"";
            string PsArguments = PsArg1 + PsArg2 + PsArg3;

            string ExeActionID = "opens powershell with parameters";
            EguibarIT.Housekeeping.AdHelper.Ocurrences Ocurrences = EguibarIT.Housekeeping.AdHelper.Ocurrences.Twelve;

            // Call the function to create task
            EguibarIT.Housekeeping.TaskScheduler.NewScheduledTask(TaskName, _serviceaccount, Description, PsArguments, ExeActionID, Ocurrences);
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

    /// <summary>
    /// <para type="synopsis">Create a new Scheduled Tasks which executes Set-PrivilegedComputerHousekeeping</para>
    /// <para type="description">Create a new Scheduled Tasks, which will trigger 12 times per day. The Action configured 
    /// is a PowerShell function from the Housekeeping module: Set-PrivilegedComputerHousekeeping</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-SchedulePrivilegedComputers -ServiceAccount "SA_R2D2" -SearchRootDN "OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-SchedulePrivilegedComputers "SA_R2D2" "OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    /// <remarks>Create new Scheduled Task which ensures that PAWs and Infrastructure Services are managed as Tier0 assets by adding those to its corresponding group.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "SchedulePrivilegedComputers", ConfirmImpact = ConfirmImpact.Low)]
    [OutputType(typeof(int))]
    public class SchedulePrivilegedComputers : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[String] SamAccountName of the group managed service account used to execute the scheduled task.</para>
        ///     <para type="description">SamAccountName of the service account (or gMSA) used to execute the scheduled task.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[String] SamAccountName of the group managed service account used to execute the scheduled task."
            )]
        [ValidateNotNullOrEmpty]
        public string ServiceAccount
        {
            get { return _serviceaccount; }
            set { _serviceaccount = value; }
        }

        private string _serviceaccount;

        /// <summary>
        ///     <para type="inputType">[STRING] Admin  OU Distinguished Name. (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        ///     <para type="description">Distinguished Name of the container where the computers are located.</para>
        /// </summary>
        [Parameter(
               Position = 1,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Admin OU Distinguished Name."
            )]
        [ValidateNotNullOrEmpty]
        public string SearchRootDN
        {
            get { return _searchrootdn; }
            set { _searchrootdn = value; }
        }

        private string _searchrootdn;

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

            // Set variables
            string TaskName = "PAWs and Infra. Servers";
            string Description = "Privileged Computers housekeeping. PAWs and Infrastructure Services must be managed by Tier0 administrators and services. Each type must belong to a certain group(SL_PAWS or SL_InfrastructureServices) in order to be mantained. This task will make sure the computer objects are members of the corresponding groups (for example, if a new computer object is created within PAW or InfraServers container, this object will automatically be added to the corresponding group based on OS; to SL_PAWs if workstation or to SL_InfrastructureServers if server).";

            string PsArg1 = "-Command \" & { Import-Module EguibarIT.Housekeeping; ";
            string PsArg2 = string.Format("Set-PrivilegedComputerHousekeeping -SearchRootDN \"{0}\"; ", _searchrootdn);
            string PsArg3 = "exit $LASTEXITCODE }\"";
            string PsArguments = PsArg1 + PsArg2 + PsArg3;

            string ExeActionID = "opens powershell with parameters";
            EguibarIT.Housekeeping.AdHelper.Ocurrences Ocurrences = EguibarIT.Housekeeping.AdHelper.Ocurrences.Twelve;

            // Call the function to create task
            EguibarIT.Housekeeping.TaskScheduler.NewScheduledTask(TaskName, _serviceaccount, Description, PsArguments, ExeActionID, Ocurrences);
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

    /// <summary>
    /// <para type="synopsis">Create a new Scheduled Tasks which executes Set-PrivilegedGroupsHousekeeping</para>
    /// <para type="description">Create a new Scheduled Tasks, which will trigger 12 times per day. The Action configured 
    /// is a PowerShell function from the Housekeeping module: Set-PrivilegedGroupsHousekeeping</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters.</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-SchedulePrivilegedGroups -ServiceAccount "SA_R2D2" -AdminUsersDN "OU=Users,OU=Admin,DC=EguibarIT,DC=local" -ExcludeList "Administrator", "TheGood"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet.</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-SchedulePrivilegedGroups "SA_R2D2" "OU=Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    /// <remarks>Create new Scheduled Task which ensures only Privileged and Semi-Privileged users are members of Admin groups.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "SchedulePrivilegedGroups", ConfirmImpact = ConfirmImpact.Low)]
    [OutputType(typeof(int))]
    public class SchedulePrivilegedGroups : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[String] SamAccountName of the group managed service account used to execute the scheduled task.</para>
        ///     <para type="description">SamAccountName of the service account (or gMSA) used to execute the scheduled task.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[String] SamAccountName of the group managed service account used to execute the scheduled task."
            )]
        [ValidateNotNullOrEmpty]
        public string ServiceAccount
        {
            get { return _serviceaccount; }
            set { _serviceaccount = value; }
        }

        private string _serviceaccount;

        /// <summary>
        ///     <para type="inputType">[STRING] Admin Users OU Distinguished Name. (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        ///     <para type="description">Distinguished Name of the container where the users are located.</para>
        /// </summary>
        [Parameter(
               Position = 1,
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
        ///     <para type="inputType">List[STRING] user list array</para>
        ///     <para type="description">Userlist to be excluded from this process.</para>
        /// </summary>
        [Parameter(
               Position = 2,
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

            // Set variables
            string TaskName = "Privileged Groups";
            string Description = "Privileged Groups housekeeping. Ensure that any group stored on Admin OU(Tier0) contains only authorized users. An Authorized user is a user created and mantained by the Tier0 Admins, and is usually identified by the SamAccountName suffix of either _T0, or _T1 or _T2. Any non authorized user will be inmediatelly revoked from these groups.";

            string PsArg1 = "-Command \" & { Import-Module EguibarIT.Housekeeping; ";
            string PsArg2;
            if (_excludelist == null)
            {
                PsArg2 = string.Format("Set-PrivilegedGroupsHousekeeping -AdminUsersDN \"{0}\"; ", AdminUsersDN);
            }
            else
            {
                PsArg2 = string.Format("Set-PrivilegedGroupsHousekeeping -AdminUsersDN \"{0}\" -ExcludeList {1}; ", AdminUsersDN, _excludelist);
            }

            string PsArg3 = "exit $LASTEXITCODE }\"";
            string PsArguments = PsArg1 + PsArg2 + PsArg3;

            string ExeActionID = "opens powershell with parameters";
            EguibarIT.Housekeeping.AdHelper.Ocurrences Ocurrences = EguibarIT.Housekeeping.AdHelper.Ocurrences.Twelve;

            // Call the function to create task
            EguibarIT.Housekeeping.TaskScheduler.NewScheduledTask(TaskName, _serviceaccount, Description, PsArguments, ExeActionID, Ocurrences);
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

    /// <summary>
    /// <para type="synopsis">Create a new Scheduled Tasks which executes Set-SemiPrivilegedKeyPairCheck</para>
    /// <para type="description">Create a new Scheduled Tasks, which will trigger 12 times per day. The Action configured 
    /// is a PowerShell function from the Housekeeping module: Set-SemiPrivilegedKeyPairCheck</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters. </para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleSemiPrivilegedKeyPair -ServiceAccount "SA_R2D2" -AdminUsersDN "OU=Users,OU=Admin,DC=EguibarIT,DC=local" -ExcludeList "Administrator", "TheGood"</code>
    ///         </para>
    ///    </example>
    ///    <example>
    ///         <para>This example shows how to use this CMDlet. </para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleSemiPrivilegedKeyPair "SA_R2D2" "OU=Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    /// <remarks>Create new Scheduled Task which ensures only each Semi-Privileged users does have a valid Non-Privileged user.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "ScheduleSemiPrivilegedKeyPair", ConfirmImpact = ConfirmImpact.Low)]
    [OutputType(typeof(int))]
    public class ScheduleSemiPrivilegedKeyPair : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[String] SamAccountName of the group managed service account used to execute the scheduled task.</para>
        ///     <para type="description">SamAccountName of the service account (or gMSA) used to execute the scheduled task.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[String] SamAccountName of the group managed service account used to execute the scheduled task."
            )]
        [ValidateNotNullOrEmpty]
        public string ServiceAccount
        {
            get { return _serviceaccount; }
            set { _serviceaccount = value; }
        }

        private string _serviceaccount;

        /// <summary>
        ///     <para type="inputType">[STRING] Admin User Account OU Distinguished Name. (ej. OU=Users,OU=Admin,DC=EguibarIT,DC=local).</para>
        ///     <para type="description">Distinguished Name of the container where the users are located.</para>
        /// </summary>
        [Parameter(
               Position = 1,
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
        ///     <para type="inputType">List[STRING] user list array</para>
        ///     <para type="description">Userlist to be excluded from this process.</para>
        /// </summary>
        [Parameter(
               Position = 2,
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

            // Set variables
            string TaskName = "Integrity check of Semi-Privileged accounts";
            string Description = "A Semi-Privileged account can only exist if a Non-Privileged account does. \n\n";
            Description += "If the key-pair non-privileged account gets disabled, the semi-privileged account must be disabled as well. \n\n";
            Description += "If the key-pair non-privileged account does not exist, the Semi-Privileged account must be deleted inmediatelly. \n\n";
            Description += "This function will read all Semi-Privileged accounts and will search its key-pair (SID of the non privileged account), applying the mentioned actions.";

            string PsArg1 = "-Command \" & { Import-Module EguibarIT.Housekeeping; ";
            string PsArg2;
            if (_excludelist == null)
            {
                PsArg2 = string.Format("Set-SemiPrivilegedKeyPairCheck -AdminUsersDN \"{0}\"; ", AdminUsersDN);
            }
            else
            {
                PsArg2 = string.Format("Set-SemiPrivilegedKeyPairCheck -AdminUsersDN \"{0}\" -ExcludeList {1}; ", AdminUsersDN, string.Join(",", _excludelist.ToArray()));
            }

            string PsArg3 = "exit $LASTEXITCODE }\"";
            string PsArguments = PsArg1 + PsArg2 + PsArg3;

            string ExeActionID = "opens powershell with parameters";
            EguibarIT.Housekeeping.AdHelper.Ocurrences Ocurrences = EguibarIT.Housekeeping.AdHelper.Ocurrences.Twelve;

            // Call the function to create task
            EguibarIT.Housekeeping.TaskScheduler.NewScheduledTask(TaskName, _serviceaccount, Description, PsArguments, ExeActionID, Ocurrences);
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

    /// <summary>
    /// <para type="synopsis">Create a new Scheduled Tasks which executes Set-ServiceAccountHousekeeping</para>
    /// <para type="description">Create a new Scheduled Tasks, which will trigger 12 times per day. The Action configured 
    /// is a PowerShell function from the Housekeeping module: Set-ServiceAccountHousekeeping</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleServiceAccounts -ServiceAccount "SA_R2D2" -ServiceAccountDN "OU=Service Accounts,OU=Admin,DC=EguibarIT,DC=local" -SAGroupName "SG_T0SA"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleServiceAccounts "SA_R2D2" "OU=Service Accounts,OU=Admin,DC=EguibarIT,DC=local" "SG_T0SA"</code>
    ///         </para>
    ///     </example>
    /// <remarks>Create new Scheduled Task which ensure each Service Account is member of the corresponding tier group.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "ScheduleServiceAccounts", ConfirmImpact = ConfirmImpact.Low)]
    [OutputType(typeof(int))]
    public class ScheduleServiceAccounts : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[STRING] SamAccountName of the group managed service account used to execute the scheduled task.</para>
        ///     <para type="description">SamAccountName of the service account (or gMSA) used to execute the scheduled task.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[String] SamAccountName of the group managed service account used to execute the scheduled task."
            )]
        [ValidateNotNullOrEmpty]
        public string ServiceAccount
        {
            get { return _serviceaccount; }
            set { _serviceaccount = value; }
        }

        private string _serviceaccount;

        /// <summary>
        ///     <para type="inputType">[STRING] representing the Distinguished Name of the object (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        ///     <para type="description">Distinguished Name of the container where the Service Accounts are located.</para>
        /// </summary>
        [Parameter(
               Position = 1,
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
        ///     <para type="inputType">[STRING] Name of the corresponding tier service account group (For tier0: SG_T0SA; for Tier1: SG_T1SA; for Tier2: SG_T2SA)</para>
        ///     <para type="description">Name of the corresponding tier service account group (For tier0: SG_T0SA; for Tier1: SG_T1SA; for Tier2: SG_T2SA).</para>
        /// </summary>
        [Parameter(
               Position = 2,
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

            string TaskName = null;
            string Description = "For each of the service accounts stored within the SA container, make it part of the corresponding service account group based on its tier. For example, the container of Tier0 service accounts must belong to the group SG_T0SA.";

            switch (_sagroupname)
            {
                case "SG_T0SA":
                    TaskName = "T0 ServiceAccount housekeeping";
                    break;

                case "SG_T1SA":
                    TaskName = "T1 ServiceAccount housekeeping";
                    break;

                case "SG_T2SA":
                    TaskName = "T2 ServiceAccount housekeeping";
                    break;
            }

            // Set variables
            string PsArg1 = "-Command \" & { Import-Module EguibarIT.Housekeeping; ";
            string PsArg2 = string.Format("Set-ServiceAccountHousekeeping -ServiceAccountDN \"{0}\" -SAGroupName {1}; ", _serviceaccountdn, _sagroupname);
            string PsArg3 = "exit $LASTEXITCODE }\"";
            string PsArguments = PsArg1 + PsArg2 + PsArg3;

            string ExeActionID = "opens powershell with parameters";
            EguibarIT.Housekeeping.AdHelper.Ocurrences Ocurrences = EguibarIT.Housekeeping.AdHelper.Ocurrences.Twelve;

            // Call the function to create task
            EguibarIT.Housekeeping.TaskScheduler.NewScheduledTask(TaskName, _serviceaccount, Description, PsArguments, ExeActionID, Ocurrences);
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

    /// <summary>
    /// <para type="synopsis">Create a new Scheduled Tasks which executes on a given event</para>
    /// <para type="description">Create a new Scheduled Tasks, which will trigger whenever an EventID is generated</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters calling a PowerShell script</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleEventTask -ServiceAccount 'gMSA_AdTskSchd' -PsArguments '-File C:\PsScripts\Hello.ps1' -EventID 4624 -Logname 'Security'</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters calling a PowerShell command</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleEventTask -ServiceAccount 'gMSA_AdTskSchd' -PsArguments '-Command "& { Import-Module EguibarIT.Housekeeping; Get-RandomPassword -PasswordLength 14 -PasswordComplexity VeryHigh; exit $LASTEXITCODE }"' -EventID 4624 -Logname 'Security'</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-ScheduleEventTask</code>
    ///         </para>
    ///     </example>
    /// <remarks>Create new Scheduled Task which triggers when specific EventID is generated.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "ScheduleEventTask", ConfirmImpact = ConfirmImpact.Low)]
    [OutputType(typeof(int))]
    public class ScheduleEventTask : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[STRING] SamAccountName of the group managed service account used to execute the scheduled task.</para>
        ///     <para type="description">SamAccountName of the service account (or gMSA) used to execute the scheduled task.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[String] SamAccountName of the group managed service account used to execute the scheduled task."
            )]
        [ValidateNotNullOrEmpty]
        public string ServiceAccount
        {
            get { return _serviceaccount; }
            set { _serviceaccount = value; }
        }
        private string _serviceaccount;



        /// <summary>
        ///     <para type="inputType">[STRING] PsArguments provides the arguments for PowerShell. A command or a PS! file can be parsed</para>
        ///     <para type="description">Arguments for PowerShell to execute.</para>
        /// </summary>
        [Parameter(
               Position = 1,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[STRING] PsArguments provides the arguments for PowerShell. A command or a PS! file can be parsed"
            )]
        [ValidateNotNullOrEmpty]
        public string PsArguments
        {
            get { return _psArguments; }
            set { _psArguments = value; }
        }
        private string _psArguments;



        /// <summary>
        ///     <para type="inputType">[STRING] EventID that will trigger this task</para>
        ///     <para type="description">EventID that will trigger this task</para>
        /// </summary>
        [Parameter(
               Position = 2,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[STRING] EventID that will trigger this task"
            )]
        [ValidateNotNullOrEmpty]
        public string EventID
        {
            get { return _eventid; }
            set { _eventid = value; }
        }
        private string _eventid;



        /// <summary>
        ///     <para type="inputType">[STRING] Logname is the name of the log file where the EventID is going to be generated.</para>
        ///     <para type="description">Name of the log file</para>
        /// </summary>
        [Parameter(
               Position = 3,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "[STRING] Logname is the name of the log file where the EventID is going to be generated."
            )]
        [ValidateNotNullOrEmpty]
        public string Logname
        {
            get { return _logname; }
            set { _logname = value; }
        }
        private string _logname;


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

            string TaskName = string.Format("EventID {0} trigger", EventID);
            string Description = string.Format("Scheduled task that will automatically trigger when eventID {0} is generated within {1} log file. This task will call Powershell to execute \"{2}\"", EventID, Logname, PsArguments);

            Dictionary<string, string> dicValueQueries = new Dictionary<string, string>();
            dicValueQueries.Add("eventChannel", "Event/System/Channel");
            dicValueQueries.Add("eventRecordID", "Event/System/EventRecordID");
            dicValueQueries.Add("eventSeverity", "Event/System/Level");
            dicValueQueries.Add("eventData", "Event/EventData/Data");

            // Set variables
            //string PsArg1 = "-Command \" & { Import-Module EguibarIT.Housekeeping; ";
            //string PsArg2 = string.Format("Set-ServiceAccountHousekeeping -ServiceAccountDN \"{0}\" -SAGroupName {1}; ", _serviceaccountdn, _sagroupname);
            //string PsArg3 = "exit $LASTEXITCODE }\"";
            //string PsArguments = PsArg1 + PsArg2 + PsArg3;

            string ExeActionID = "opens powershell with parameters";

            try
            {
                // Call the function to create task
                EguibarIT.Housekeeping.TaskScheduler.NewScheduledEventTask(TaskName, ServiceAccount, Description, PsArguments, ExeActionID, dicValueQueries, EventID, Logname);
            }//end try
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
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

}//end namespace