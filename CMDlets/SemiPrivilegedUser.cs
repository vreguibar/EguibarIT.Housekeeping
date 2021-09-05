using System;
using System.Collections.ObjectModel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.IO;
using System.Management.Automation;
using System.Text;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    ///     <para type="synopsis">Create or modify a Semi-Privileged account.</para>
    ///     <para type="description">Create or modify a Semi-Privileged account.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>New-SemiPrivilegedUser dvader dvader@eguibarit.com T0 'OU=Admin Accounts,OU=Admin,DC=EguibarIT,DC=local' 'DelegationModel@eguibarit.com' 'smtpout.europe.secureserver.net'</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>New-SemiPrivilegedUser -SamAccountName dvader -EmailTo dvader@eguibarit.com -AccountType T0 -AdminUsersDN 'OU=Admin Accounts,OU=Admin,DC=EguibarIT,DC=local' -From 'DelegationModel@eguibarit.com' -SMTPserver 'smtpout.europe.secureserver.net'</code>
    ///         </para>
    ///     </example>
    ///     <remarks>Create or modify a Semi-Privileged account.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.New, "SemiPrivilegedUser", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(int))]
    public class SemiPrivilegedUser : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[STRING] representing the user SamAccountName.</para>
        /// <para type="description">Identity of the user getting the new Admin Account (Semi-Privileged user).</para>
        /// </summary>
        [Parameter(
            Position = 0,
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Identity of the user getting the new Admin Account (Semi-Privileged user)."
            )]
        [ValidateNotNullOrEmpty]
        public string SamAccountName
        {
            get { return _samaccountname; }
            set { _samaccountname = value; }
        }

        private string _samaccountname;

        /// <summary>
        /// <para type="inputType">[STRING] representing the valid E-mail address.</para>
        /// <para type="description">Valid Email of the target user. This address will be used to send information to her/him.</para>
        /// </summary>
        [Parameter(
            Position = 1,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Valid Email of the target user. This address will be used to send information to her/him."
            )]
        [ValidateNotNullOrEmpty]
        [ValidatePattern(@"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$")]
        public string EmailTo
        {
            get { return _emailto; }
            set { _emailto = value; }
        }

        private string _emailto;

        /// <summary>
        /// <para type="inputType">[STRING] Indicating which kind of Semi-Privileged account to be created. Valid values are T0 or T1 or T2</para>
        /// <para type="description">Must specify the account type. Valid values are T0 or T1 or T2</para>
        /// </summary>
        [Parameter(
            Position = 2,
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Must specify the account type. Valid values are T0 or T1 or T2"
            )]
        [ValidateSet("T0", "T1", "T2", IgnoreCase = true)]
        public string AccountType
        {
            get { return _accounttype; }
            set { _accounttype = value; }
        }

        private string _accounttype;

        /// <summary>
        /// <para type="inputType">[STRING] Admin User Account OU Distinguished Name. (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        /// <para type="description">Distinguished Name of the container where the Admin users are located.</para>
        /// </summary>
        [Parameter(
            Position = 3,
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
        /// <para type="inputType">[STRING] representing the valid E-mail address of sender.</para>
        /// <para type="description">Valid Email of the sending user. This address will be used to send the information and for authenticate to the SMTP server.</para>
        /// </summary>
        [Parameter(
            Position = 4,
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
            Position = 5,
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
            Position = 6,
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
            Position = 7,
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
        /// <para type="inputType">[int] representing the SMTP port number.</para>
        /// <para type="description">SMTP port number.</para>
        /// </summary>
        [Parameter(
            Position = 8,
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
        /// <para type="inputType">[STRING] representing the body template file.</para>
        /// <para type="description">Path to the body template file.</para>
        /// </summary>
        [Parameter(
            Position = 9,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Path to the body template file."
            )]
        [ValidateNotNullOrEmpty]
        public string BodyTemplate
        {
            get { return _bodytemplate; }
            set { _bodytemplate = value; }
        }

        private string _bodytemplate;

        /// <summary>
        /// <para type="inputType">[STRING] representing the attached image for the body template.</para>
        /// <para type="description">Path to the attached image of body template.</para>
        /// </summary>
        [Parameter(
            Position = 10,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Path to the attached image of body template."
            )]
        [ValidateNotNullOrEmpty]
        public string BodyImage
        {
            get { return _bodyimage; }
            set { _bodyimage = value; }
        }

        private string _bodyimage;

        /// <summary>
        /// <para type="inputType">[STRING] representing the Password body template file.</para>
        /// <para type="description">Path to the Password body template file.</para>
        /// </summary>
        [Parameter(
            Position = 11,
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

            bool sendEmail = false;
            bool sendPassword = false;

            String _newSamAccountName = null;
            String NewPWD = EguibarIT.Housekeeping.Helpers.RandomPassword(true, true, true, false, false, 16);

            if (_emailto != null) { sendEmail = true; }

            // set up domain context
            using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain))
            {
                // Get the User Principal
                ExtPrincipal.UserPrincipalEx user = ExtPrincipal.UserPrincipalEx.FindByIdentity(ctx, _samaccountname);

                // Check if user IS a UserPrincipal (User exists)
                if (user == null)
                {
                    WriteWarning("[WARNING] Standard User does not exist. Before creating a Semi-Privileged user, a standard user (Non-Privileged user) must exist. Make sure a standard user exist before proceeding.");

                    sendEmail = false;
                    sendPassword = false;
                }
                else
                {
                    WriteVerbose("[PROCESS] Standard user (Non-Privileged user) found! Proceeding to create the Semi-Privileged user.");
                    try
                    {
                        using (PrincipalContext pctx = new PrincipalContext(ContextType.Domain,
                                                                        EguibarIT.Housekeeping.AdHelper.AdDomain.GetAdFQDN(),
                                                                        _adminusersdn))
                        {
                            _newSamAccountName = string.Format("{0}_{1}", user.SamAccountName, _accounttype.ToUpper());

                            string _Surename = null;
                            try
                            {
                                _Surename = user.Surname.ToString();
                            }
                            catch { }

                            string _GivenName = null;
                            try
                            {
                                _GivenName = user.GivenName.ToString();
                            }
                            catch { }

                            //Create new Semi-Privileged user
                            using (ExtPrincipal.UserPrincipalEx newUser = new ExtPrincipal.UserPrincipalEx(pctx))
                            {
                                newUser.SamAccountName = string.Format("{0}_{1}", user.SamAccountName, _accounttype.ToUpper());
                                newUser.UserPrincipalName = string.Format("{0}_{1}@{2}", user.SamAccountName, _accounttype.ToUpper(), EguibarIT.Housekeeping.AdHelper.AdDomain.GetAdFQDN());

                                if (_Surename == null)
                                {
                                    newUser.Name = string.Format("{0} ({1})", _GivenName, _accounttype.ToUpper());
                                    newUser.DisplayName = string.Format("{0} ({1})", _GivenName, _accounttype.ToUpper());
                                    newUser.GivenName = _GivenName;
                                }
                                else
                                {
                                    newUser.Name = string.Format("{0}, {1} ({2})", _Surename.ToUpper(), _GivenName, _accounttype.ToUpper());
                                    newUser.DisplayName = string.Format("{0}, {1} ({2})", _Surename.ToUpper(), _GivenName, _accounttype.ToUpper());
                                }

                                if (_GivenName == null)
                                {
                                    newUser.Name = string.Format("{0} ({1})", _Surename.ToUpper(), _accounttype.ToUpper());
                                    newUser.DisplayName = string.Format("{0} ({1})", _Surename.ToUpper(), _accounttype.ToUpper());
                                    newUser.Surname = _Surename.ToUpper();
                                }

                                //newUser.DisplayName = string.Format("{0}, {1} ({2})", user.Surname.ToUpper(), user.GivenName, _accounttype.ToUpper());

                                newUser.Enabled = true;
                                newUser.ScriptPath = null;
                                newUser.HomeDirectory = null;
                                newUser.HomeDrive = null;
                                newUser.EmailAddress = user.EmailAddress;
                                newUser.EmployeeId = user.EmployeeId;
                                newUser.employeeNumber = user.Sid.Value.ToString();
                                newUser.City = user.City;
                                newUser.Company = user.Company;
                                newUser.CountryName = user.CountryName;
                                newUser.Department = user.Department;
                                newUser.Division = user.Division;
                                newUser.MobilePhone = user.MobilePhone;
                                newUser.Office = user.Office;
                                newUser.TelephoneNumber = user.TelephoneNumber;
                                newUser.Organization = user.Organization;
                                newUser.MiddleName = user.MiddleName;
                                newUser.POBox = user.POBox;
                                newUser.PostalCode = user.PostalCode;
                                newUser.State = user.State;
                                newUser.StreetAddress = user.StreetAddress;
                                newUser.Title = user.Title;
                                newUser.EmployeeType = _accounttype.ToUpper();
                                newUser.AllowDialin = false;
                                newUser.Description = string.Format("{0} Semi-Privileged Account", _accounttype.ToUpper());
                                newUser.SupportedEncryptionTypes = 24;

                                //Set the new Password
                                newUser.SetPassword(NewPWD);

                                // Change PWD at next logon
                                newUser.ExpirePasswordNow();

                                newUser.Save();

                                //The user MUST be saved before changing UserAccountControl flags, otherwise throws an error.
                                newUser.UserAccountControl = (int)(EguibarIT.Housekeeping.AdHelper.UserAccountControl.NORMAL_ACCOUNT |
                                                              EguibarIT.Housekeeping.AdHelper.UserAccountControl.NOT_DELEGATED);
                                newUser.Save();

                                // If no email address provided, try to get it from Non-Privileged user
                                if (_emailto == null)
                                {
                                    if (user.EmailAddress == null)
                                    {
                                        _emailto = user.EmailAddress.ToString();
                                    }
                                }

                                /*
                                //Change from UserPrincipal to DirectoryEntry
                                if (newUser.GetUnderlyingObjectType() == typeof(DirectoryEntry))
                                {
                                    using (var entry = (DirectoryEntry)newUser.GetUnderlyingObject())
                                    {
                                        entry.Properties["userAccountControl"][0] = (
                                            EguibarIT.Housekeeping.AdHelper.UserAccountControl.NORMAL_ACCOUNT |
                                            EguibarIT.Housekeeping.AdHelper.UserAccountControl.NOT_DELEGATED);

                                        entry.CommitChanges();
                                    }
                                }
                                */
                                /*
                                thumbnailPhoto
                                */
                            } // end using
                            if (_emailto != null)
                            {
                                WriteVerbose("[PROCESS] Semi-Privileged user created. Preparing notification email.");
                                sendEmail = true;
                                sendPassword = true;
                            }
                            else
                            {
                                WriteVerbose("[PROCESS] Semi-Privileged user created. No EmailTo was provided, so no notification will be created. Password must be set again manually.");
                                sendEmail = false;
                                sendPassword = false;
                            }
                        }// end using
                    } //end try
                    catch (Exception Ex)
                    {
                        if (Ex.Message == "The object already exists.\r\n")
                        {
                            WriteVerbose("[PROCESS] The Semi-Privileged user already exists. Modifying the account.");
                            try
                            {
                                _newSamAccountName = string.Format("{0}_{1}", user.SamAccountName, _accounttype.ToUpper());

                                string _Surename = null;
                                try
                                {
                                    _Surename = user.Surname.ToString();
                                }
                                catch { }

                                string _GivenName = null;
                                try
                                {
                                    _GivenName = user.GivenName.ToString();
                                }
                                catch { }

                                using (ExtPrincipal.UserPrincipalEx existingAdmin = ExtPrincipal.UserPrincipalEx.FindByIdentity(ctx, IdentityType.SamAccountName, _newSamAccountName))
                                {
                                    existingAdmin.UserPrincipalName = string.Format("{0}_{1}@{2}", user.SamAccountName, _accounttype.ToUpper(), EguibarIT.Housekeeping.AdHelper.AdDomain.GetAdFQDN());
                                    if (_Surename == null)
                                    {
                                        //existingAdmin.Name = string.Format("{0} ({1})", _GivenName, _accounttype.ToUpper());
                                        existingAdmin.DisplayName = string.Format("{0} ({1})", _GivenName, _accounttype.ToUpper());
                                        existingAdmin.GivenName = _GivenName;
                                    }
                                    else
                                    {
                                        //existingAdmin.Name = string.Format("{0}, {1} ({2})", _Surename.ToUpper(), _GivenName, _accounttype.ToUpper());
                                        existingAdmin.DisplayName = string.Format("{0}, {1} ({2})", _Surename.ToUpper(), _GivenName, _accounttype.ToUpper());
                                    }

                                    if (_GivenName == null)
                                    {
                                        //existingAdmin.Name = string.Format("{0} ({1})", _Surename.ToUpper(), _accounttype.ToUpper());
                                        existingAdmin.DisplayName = string.Format("{0} ({1})", _Surename.ToUpper(), _accounttype.ToUpper());
                                        existingAdmin.Surname = _Surename.ToUpper();
                                    }

                                    ////existingAdmin.Name = string.Format("{0}, {1} ({2})", user.Surname.ToUpper(), user.GivenName, _accounttype.ToUpper());
                                    //existingAdmin.DisplayName = string.Format("{0}, {1} ({2})", user.Surname.ToUpper(), user.GivenName, _accounttype.ToUpper());
                                    //existingAdmin.Surname = user.Surname.ToUpper();
                                    //existingAdmin.GivenName = user.GivenName;
                                    existingAdmin.Enabled = true;
                                    existingAdmin.ScriptPath = null;
                                    existingAdmin.HomeDirectory = null;
                                    existingAdmin.HomeDrive = null;
                                    existingAdmin.EmailAddress = user.EmailAddress;
                                    existingAdmin.EmployeeId = user.EmployeeId;
                                    existingAdmin.employeeNumber = user.Sid.Value.ToString();
                                    existingAdmin.City = user.City;
                                    existingAdmin.Company = user.Company;
                                    existingAdmin.CountryName = user.CountryName;
                                    existingAdmin.Department = user.Department;
                                    existingAdmin.Division = user.Division;
                                    existingAdmin.MobilePhone = user.MobilePhone;
                                    existingAdmin.Office = user.Office;
                                    existingAdmin.TelephoneNumber = user.TelephoneNumber;
                                    existingAdmin.Organization = user.Organization;
                                    existingAdmin.MiddleName = user.MiddleName;
                                    existingAdmin.POBox = user.POBox;
                                    existingAdmin.PostalCode = user.PostalCode;
                                    existingAdmin.State = user.State;
                                    existingAdmin.StreetAddress = user.StreetAddress;
                                    existingAdmin.Title = user.Title;
                                    existingAdmin.EmployeeType = _accounttype.ToUpper();
                                    existingAdmin.AllowDialin = false;
                                    existingAdmin.Description = string.Format("{0} Semi-Privileged Account", _accounttype.ToUpper());
                                    existingAdmin.SupportedEncryptionTypes = 24;

                                    existingAdmin.Save();

                                    existingAdmin.UserAccountControl = (int)(EguibarIT.Housekeeping.AdHelper.UserAccountControl.NORMAL_ACCOUNT |
                                                                              EguibarIT.Housekeeping.AdHelper.UserAccountControl.NOT_DELEGATED);
                                    existingAdmin.Save();

                                    if (existingAdmin.GetUnderlyingObjectType() == typeof(DirectoryEntry))
                                    {
                                        using (var entry = (DirectoryEntry)existingAdmin.GetUnderlyingObject())
                                        {
                                            entry.MoveTo(new DirectoryEntry(string.Format("LDAP://{0}", _adminusersdn)));
                                        }
                                    }//end if
                                }//end using UserPrincipalEx

                                sendEmail = true;
                                sendPassword = false;

                                WriteVerbose("[PROCESS] Existing Semi-Privileged user was modified succesfully.");
                            }//end try
                            catch (Exception Ex1)
                            {
                                sendEmail = false;
                                sendPassword = false;

                                WriteWarning("Something went wrong while modifying existing Semi-Privileged user.");
                                WriteWarning(Ex1.ToString());
                            }
                        }//end if
                    }//end catch
                }//end else
            }//end using

            //Preparing email notifications
            //only if EmailTo is provided
            if (_emailto != null)
            {
                string body;

                //Send notification Email
                try
                {
                    if (sendEmail)
                    {
                        WriteVerbose("[PROCESS] Preparing the notification email...");
                        bool _useDefaultBody = false;
                        Bitmap bmp;
                        MailAttachment attach = null;

                        // Create empty stream
                        MemoryStream _stream = new MemoryStream();

                        try
                        {
                            //Check if body template was passed
                            if (_bodytemplate != null)
                            {
                                // Use the template passed from parameter
                                body = System.IO.File.ReadAllText(@_bodytemplate);
                            }
                            else
                            {
                                //Body not present. Use default one
                                body = HTML._newID;
                                _useDefaultBody = true;
                                WriteVerbose("[PROCESS] No body template was parsed. Using default body template instead.");
                            }

                            // Get the mail message and replace strings.
                            body = body.Replace("#@UserID@#", _newSamAccountName);
                            body = body.Replace("#@DomainName@#", EguibarIT.Housekeeping.AdHelper.AdDomain.GetNetbiosDomainName());

                            // If no body is passed, use the default
                            if (_useDefaultBody)
                            {
                                ///////////////
                                // Get image from Resources
                                bmp = new Bitmap(EguibarIT.Housekeeping.Resources.Resource1.Picture1);

                                // Save the Bitmap to the stream
                                bmp.Save(_stream, System.Drawing.Imaging.ImageFormat.Bmp);
                                //Set the string to the start
                                _stream.Position = 0;

                                //Build the attachment
                                attach = new MailAttachment(_stream, "Picture1.jpg");
                                bmp.Dispose();
                            }

                            // If boody and image is passed, attach image
                            if (!_useDefaultBody)
                            {
                                if (_bodyimage != null)
                                {
                                    bmp = new Bitmap(@_bodyimage);

                                    // Save the Bitmap to the stream
                                    bmp.Save(_stream, System.Drawing.Imaging.ImageFormat.Bmp);
                                    //Set the string to the start
                                    _stream.Position = 0;

                                    //Build the attachment
                                    attach = new MailAttachment(_stream, Path.GetFileName(@_bodyimage));
                                    bmp.Dispose();
                                }
                            }
                            //_stream.Dispose();
                            try
                            {
                                // Send the email
                                Email.NewEmail(_emailto,
                                               body,
                                               "New Semi-Privileged account based on the AD Delegation Model",
                                               _from,
                                               "Delegation Model toolset",
                                               _credentialUser,
                                               _credentialPassword,
                                               _smtphost,
                                               _smtpport,
                                               attach);

                                WriteVerbose("[PROCESS] Notification email sent succesfully.");
                            }//end try
                            catch (Exception ExEmail)
                            {
                                WriteWarning("Notification email could not be sent.");
                                WriteWarning(ExEmail.ToString());
                            }
                        }//end try
                        catch (Exception Ex)
                        {
                            sendPassword = false;

                            WriteWarning("Notification email for newly created user could not be sent.");
                            WriteWarning(Ex.ToString());
                        }

                        _stream.Dispose();

                    }//end if sendEmail
                }//end try
                catch (Exception Ex)
                {
                    WriteWarning("Email notification could not be sent.");
                    WriteWarning(Ex.ToString());
                    sendPassword = false;
                }

                //Send encrypted Email PWD
                try
                {
                    if (sendPassword)
                    {
                        WriteVerbose("[PROCESS] Preparing the encrypted password email...");
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
                        catch (Exception Ex)
                        {
                            WriteWarning("Encrypted email containing password could not be sent.");
                            WriteWarning(Ex.ToString());
                        }
                    }//end if sendPassword
                }//end try
                catch (Exception Ex)
                {
                    WriteWarning("Encrypted email containing password could not be sent.");
                    WriteWarning(Ex.ToString());
                }
            }//end if emailTo
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
    } //END class Xxxxx
} // END namespace EguibarIT.Delegation