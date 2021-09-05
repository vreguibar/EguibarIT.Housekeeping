using System;
using System.Collections.ObjectModel;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Management.Automation;
using System.Security.Principal;
using System.Text;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Reset password of semi-privileged user</para>
    /// <para type="description">Reset password of semi-privileged user</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-SemiPrivilegedUsersPwdReset dvader_t0</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Set-SemiPrivilegedUsersPwdReset -SamAccountName dvader_t0</code>
    ///         </para>
    /// </example>
    /// <remarks>Reset password of semi-privileged user</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "SemiPrivilegedUsersPwdReset", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(ExtPrincipal.UserPrincipalEx))]
    public class SemiPrivilegedUsersPwdReset : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[STRING] representing the user SamAccountName.</para>
        /// <para type="description">Identity of the Semi-Privileged user which will have a new password.</para>
        /// </summary>
        [Parameter(
            Position = 0,
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Identity of the Semi-Privileged user which will have a new password."
            )]
        [ValidateNotNullOrEmpty]
        public string SamAccountName
        {
            get { return _samaccountname; }
            set { _samaccountname = value; }
        }

        private string _samaccountname;

        /// <summary>
        /// <para type="inputType">[STRING] representing the valid E-mail address of sender.</para>
        /// <para type="description">Valid Email of the sending user. This address will be used to send the information and for authenticate to the SMTP server.</para>
        /// </summary>
        [Parameter(
            Position = 1,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Valid Email of the sending user. This address will be used to send the information and for authenticate to the SMTP server."
            )]
        [ValidateNotNullOrEmpty]
        [ValidatePattern(@"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$")]
        public string From
        {
            get { return _from; }
            set { _from = value; }
        }

        private string _from;

        /// <summary>
        /// <para type="inputType">[STRING] representing the valid user for authenticating to SMTP.</para>
        /// <para type="description">User for authenticate to the SMTP server.</para>
        /// </summary>
        [Parameter(
            Position = 2,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "User for authenticate to the SMTP server."
            )]
        [ValidateNotNullOrEmpty]
        public string CredentialUser
        {
            get { return _credentialUser; }
            set { _credentialUser = value; }
        }

        private string _credentialUser = null;

        /// <summary>
        /// <para type="inputType">[STRING] representing the valid password for authenticating to SMTP (User is E-mail address of sender).</para>
        /// <para type="description">Password for authenticate to the SMTP server.</para>
        /// </summary>
        [Parameter(
            Position = 3,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Password for authenticate to the SMTP server. (User is E-mail address of sender)"
            )]
        [ValidateNotNullOrEmpty]
        public string CredentialPassword
        {
            get { return _credentialPassword; }
            set { _credentialPassword = value; }
        }

        private string _credentialPassword = null;

        /// <summary>
        /// <para type="inputType">[STRING] representing the SMTP server.</para>
        /// <para type="description">SMTP server.</para>
        /// </summary>
        [Parameter(
            Position = 4,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "SMTP server."
            )]
        [ValidateNotNullOrEmpty]
        public string SMTPserver
        {
            get { return _smtphost; }
            set { _smtphost = value; }
        }

        private string _smtphost;

        /// <summary>
        /// <para type="inputType">[INT] representing the SMTP port number.</para>
        /// <para type="description">SMTP port number.</para>
        /// </summary>
        [Parameter(
            Position = 5,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "SMTP port number."
            )]
        [ValidateNotNullOrEmpty]
        public int SMTPport
        {
            get { return _smtpport; }
            set { _smtpport = value; }
        }

        private int _smtpport = 25;

        /// <summary>
        /// <para type="inputType">[STRING] representing the Password body template file.</para>
        /// <para type="description">Path to the Password body template file.</para>
        /// </summary>
        [Parameter(
            Position = 6,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Path to the body template file."
            )]
        [ValidateNotNullOrEmpty]
        public string PwdBodyTemplate
        {
            get { return _pwdbodytemplate; }
            set { _pwdbodytemplate = value; }
        }

        private string _pwdbodytemplate;

        /// <summary>
        /// <para type="inputType">[SWITCH] indicating if password has to be sent by encrypted email.</para>
        /// <para type="description">If present, will send the new password by encrypted email.</para>
        /// </summary>
        [Parameter(
            Position = 7,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "If present, will send the new password by encrypted email."
            )]
        [ValidateNotNullOrEmpty]
        public SwitchParameter SendByEmail
        {
            get { return _sendbyemail; }
            set { _sendbyemail = value; }
        }

        private bool _sendbyemail;

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

            //Flag for sending email
            bool sendPassword = false;

            // Generate new password
            String NewPWD = EguibarIT.Housekeeping.Helpers.RandomPassword(true, true, true, true, true, 16);

            String _emailto = "";

            // set up domain context
            using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain))
            {
                // Get the User Principal - SemiPrivileged account
                using (ExtPrincipal.UserPrincipalEx user = ExtPrincipal.UserPrincipalEx.FindByIdentity(ctx, _samaccountname))
                {
                    // Check if user IS a UserPrincipal (User exists)
                    if (user == null)
                    {
                        WriteWarning("[WARNING] SemiPrivileged User does not exist. Make sure the SemiPrivileged user exist before proceeding.");

                        sendPassword = false;
                    }
                    else
                    {
                        // Get hte NonPrivileged user email
                        try
                        {
                            //Convert string to SID
                            SecurityIdentifier _NonPrivilegedSid = new SecurityIdentifier(user.employeeNumber);

                            // Get the User Principal (NonPrivileged user having the SID on EmployeeNumber of the SemiPrivileged account)
                            using (ExtPrincipal.UserPrincipalEx _NonPrivilegedUser = ExtPrincipal.UserPrincipalEx.FindByIdentity(ctx, IdentityType.Sid, _NonPrivilegedSid.ToString()))
                            {
                                //ExtPrincipal.UserPrincipalEx _NonPrivilegedUser = ExtPrincipal.UserPrincipalEx.FindByIdentity(ctx, IdentityType.Sid, user.employeeNumber);
                                if (_NonPrivilegedUser == null)
                                {
                                    WriteWarning("[WARNING] Is not possible to find the Key-Pair of the given SemiPrivileged User (SemiPrivileged user must always have its Key-Pair NonPrivilegedUser). Make sure the Non-Privileged user exist before proceeding.");

                                    sendPassword = false;
                                }//end if
                                else
                                {
                                    // Get the NonPrivileged user email
                                    _emailto = _NonPrivilegedUser.EmailAddress;

                                    if (_emailto != null)
                                    {
                                        sendPassword = true;
                                        WriteVerbose(string.Format("[PROCESS] The encrypted email will be sent to: {0}.", _emailto));
                                    }
                                }//end else
                            }//end using NonPrivileged
                        }//end try
                        catch (Exception Ex)
                        {
                            sendPassword = false;

                            WriteWarning("Something went wrong while changing the password of the Semi-Privileged user.");
                            WriteWarning(Ex.ToString());
                        }//end catch
                    }//end else

                    if (sendPassword)
                    {
                        try
                        {
                            // Set the new Password and save
                            user.SetPassword(NewPWD);
                            user.Save();

                            WriteVerbose("[PROCESS] The password has been updated succesfully.");
                        }
                        catch (Exception Ex1)
                        {
                            sendPassword = false;

                            WriteWarning("Something went wrong while changing the password of the Semi-Privileged user.");
                            WriteWarning(Ex1.ToString());
                        }//end catch
                    }//end if
                }//end using SemiPrivileged account
            }//end using ctx

            if (_sendbyemail)
            {
                try
                {
                    if (sendPassword)
                    {
                        WriteVerbose("[PROCESS] Preparing the encrypted password email...");
                        string body;
                        //Check if body template was passed
                        if (_pwdbodytemplate != null)
                        {
                            // Use the template passed from parameter
                            body = System.IO.File.ReadAllText(@_pwdbodytemplate);
                        }
                        else
                        {
                            //Body not present. Use default one
                            body = HTML._newPWD;
                            WriteVerbose("No password body template was parsed. Using default password body template instead.");
                        }

                        body = body.Replace("#@unsecurePassword@#", NewPWD);
                        body = body.Replace("#@DomainName@#", EguibarIT.Housekeeping.AdHelper.AdDomain.GetNetbiosDomainName());

                        try
                        {
                            using (PowerShell powershell = PowerShell.Create())
                            {
                                StringBuilder sb = new StringBuilder();

                                if (File.Exists(@"C:\PsScripts\Send-EncryptedEmail.ps1"))
                                {
                                    // Add a script to the PowerShell object.
                                    powershell.AddCommand(@"C:\PsScripts\Send-EncryptedEmail.ps1");

                                    //Add parameters to previous command
                                    powershell.AddParameter("Emailto", _emailto);
                                    powershell.AddParameter("EmailFrom", _from);
                                    powershell.AddParameter("Body", body);
                                    powershell.AddParameter("Subject", "New Administrative account \"Confidential\"");

                                    //Execute the script.
                                    Collection<PSObject> results = powershell.Invoke();
                                    Collection<ErrorRecord> errors = powershell.Streams.Error.ReadAll();
                                    if (errors.Count > 0)
                                    {
                                        foreach (ErrorRecord error in errors)
                                        {
                                            sb.AppendLine(error.ToString());
                                        }
                                    }
                                    else
                                    {
                                        foreach (PSObject result in results)
                                        {
                                            sb.AppendLine(result.ToString());
                                        }
                                    }
                                    WriteVerbose(sb.ToString());

                                    WriteVerbose("Encrypted email containing password was sent succesfully!");
                                }//end if file.exists
                                else
                                {
                                    WriteWarning("Cannot find the PowerShell script \"Send-EncryptedEmail.ps1\" which sends the encripted Email. Make sure that the script exist in \"C:\\PsScripts\\Send-EncryptedEmail.ps1\"");
                                }
                                //}
                            }//end using
                        }//end try
                        catch (Exception Ex2)
                        {
                            WriteWarning("Encrypted email containing password could not be sent.");
                            WriteWarning(Ex2.ToString());
                        }
                    }//end if sendPassword
                }//end try
                catch (Exception Ex3)
                {
                    WriteWarning("Encrypted email containing password could not be sent.");
                    WriteWarning(Ex3.ToString());
                }
            }//end if _sendbyemail
            else
            {
                WriteObject(NewPWD);
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
}