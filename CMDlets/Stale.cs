using System;
using System.Collections.Generic;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Find stale users in domain.</para>
    /// <para type="description">A stale user is a user which has not logon to the domain in a certain period of time. 
    /// This time is calculated on the AD attribute LastLogonTimestamp. When this value is older than the 
    /// value provided on the DaysOffset, the user is considered as stale. If the parameter 
    /// DaysOffset (integer) is ommited,a default value of 45 will be used. 
    /// The return object is a list of UserPrincipal Extended.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdStaleUser 30</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdStaleUser -DaysOffset 30</code>
    ///         </para>
    ///     </example>
    /// <remarks>Find stale users in domain.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Get, "AdStaleUser", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(ExtPrincipal.UserPrincipalEx))]
    public class AdStaleUser : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[INT] Integer representing days offset to search for the stale user.</para>
        /// <para type="description">Days offset to search for the stale user.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = false,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Days offset to search for the stale user."
            )]
        [ValidateNotNullOrEmpty]
        public int DaysOffset
        {
            get { return _daysoffset; }
            set { _daysoffset = value; }
        }

        private int _daysoffset;

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

            if (_daysoffset == 0)
            {
                _daysoffset = 45;
            }

            List<ExtPrincipal.UserPrincipalEx> StaleUsers = Helpers.StaleUsers(_daysoffset);

            WriteVerbose(string.Format("The DayOffset to search for stale objects is {0}", _daysoffset));
            WriteVerbose(string.Format("Found {0} stale objects", StaleUsers.Count));

            WriteObject(StaleUsers);
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

    /// <summary>
    /// <para type="synopsis">Find stale computer in domain.</para>
    /// <para type="description">A stale computer is a computer which has not logon to the domain in a certain period of time. 
    /// This time is calculated on the AD attribute LastLogonTimestamp. When this value is older than the 
    /// value provided on the DaysOffset, the user is considered as stale. If the parameter  
    /// DaysOffset (integer) is ommited,a default value of 45 will be used. 
    /// The return object is a list of ComputerPrincipal Extended.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdStaleComputer 30</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdStaleComputer -DaysOffset 30</code>
    ///         </para>
    /// </example>
    /// <remarks>Find stale computer in domain.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Get, "AdStaleComputer", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(ExtPrincipal.ComputerPrincipalEx))]
    public class AdStaleComputer : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[INT] Integer representing days offset to search for the stale computer.</para>
        /// <para type="description">Days offset to search for the stale computer.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = false,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Days offset to search for the stale user."
            )]
        [ValidateNotNullOrEmpty]
        public int DaysOffset
        {
            get { return _daysoffset; }
            set { _daysoffset = value; }
        }

        private int _daysoffset;

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

            if (_daysoffset == 0)
            {
                _daysoffset = 45;
            }

            List<ExtPrincipal.ComputerPrincipalEx> StaleComputers = Helpers.StaleComputers(_daysoffset);

            WriteVerbose(string.Format("The DayOffset to search for stale objects is {0}", _daysoffset));
            WriteVerbose(string.Format("Found {0} stale objects", StaleComputers.Count));

            WriteObject(StaleComputers);
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