using System;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    ///     <para type="synopsis">Generate random password</para>
    ///     <para type="description">Generate random password, with defined length and complexity.</para>
    ///         <example>
    ///             <para>This example shows how to use this CMDlet using named parameters</para>
    ///             <para>-        </para>
    ///             <para>
    ///                 <code>Get-RandomPassword -PasswordLength 14 -PasswordComplexity VeryHigh</code>
    ///             </para>
    ///         </example>
    ///         <example>
    ///             <para>This example shows how to use this CMDlet</para>
    ///             <para>-        </para>
    ///             <para>
    ///                 <code>Get-RandomPassword 14 VeryHigh</code>
    ///             </para>
    ///         </example>
    ///         <example>
    ///             <para>Create a new 10 Character Password of Uppercase/Lowercase/Numbers and store as a Secure.String in Variable called $MYPASSWORD</para>
    ///             <para>-        </para>
    ///             <para>
    ///                 <code>$MYPASSWORD = CONVERTTO-SECURESTRING (Get-RandomPassword -Length 10 -Complexity High) -asplaintext -force</code>
    ///             </para>
    ///         </example>
    ///     <remarks>Generate random password with configurable parameters.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Get, "RandomPassword", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(string))]
    public class RandomPassword : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[INT] Length of the password to be generated.</para>
        /// <para type="description">Length of the password to be generated.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Length of the password to be generated."
            )]
        [ValidateNotNullOrEmpty]
        public int PasswordLength
        {
            get { return _passwordlength; }
            set { _passwordlength = value; }
        }

        private int _passwordlength;

        /// <summary>
        /// <para type="inputType">[ValidateSet] Low, Medium, High or VeryHigh</para>
        /// <para type="description">Password complexity. Low use lowercase characters. Medium uses aditional upercase characters. High uses additional number characters. VeryHigh uses additional special characters.</para>
        /// </summary>
        [Parameter(
               Position = 1,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Password complexity. Low use lowercase characters. Medium uses aditional upercase characters. High uses additional number characters. VeryHigh uses additional special characters."
            )]
        [ValidateNotNullOrEmpty]
        [ValidateSet("Low", "Medium", "High", "VeryHigh", IgnoreCase = true)]
        public string PasswordComplexity
        {
            get { return _passwordcomplexity; }
            set { _passwordcomplexity = value; }
        }

        private string _passwordcomplexity;

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

            string GeneratedPassword = null;

            try
            {
                switch (_passwordcomplexity.ToLower())
                {
                    case "low":
                        GeneratedPassword = Helpers.RandomPassword(true, false, false, false, false, _passwordlength);
                        break;

                    case "medium":
                        GeneratedPassword = Helpers.RandomPassword(true, true, false, false, false, _passwordlength);
                        break;

                    case "high":
                        GeneratedPassword = Helpers.RandomPassword(true, true, true, false, false, _passwordlength);
                        break;

                    case "veryhigh":
                        GeneratedPassword = Helpers.RandomPassword(true, true, true, true, true, _passwordlength);
                        break;

                    default:
                        GeneratedPassword = Helpers.RandomPassword(true, true, true, false, false, _passwordlength);
                        break;
                } //end switch
            }//end try
            catch { }

            Console.WriteLine("");
            Console.WriteLine("Your new random password is:");
            Console.WriteLine("----------------------------------------");
            Console.WriteLine("");
            WriteObject(GeneratedPassword);
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine(string.Format("The password was generated using {0} complexity and {1} characters long.", _passwordcomplexity, _passwordlength));
            Console.WriteLine("");
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