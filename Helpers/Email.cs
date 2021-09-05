using System;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Management.Automation;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace EguibarIT.Housekeeping
{
    /// <summary>
    /// Sending Email class
    /// </summary>
    public class Email
    {
        /// <summary>
        /// Mail with authentication
        /// </summary>
        /// <param name="to"></param>
        /// <param name="body"></param>
        /// <param name="subject"></param>
        /// <param name="fromAddress"></param>
        /// <param name="fromDisplay"></param>
        /// <param name="credentialUser"></param>
        /// <param name="credentialPassword"></param>
        /// <param name="smtphost"></param>
        /// <param name="port"></param>
        /// <param name="attachments"></param>
        public static void NewEmail(string to,
                                    string body,
                                    string subject,
                                    string fromAddress,
                                    string fromDisplay,
                                    string credentialUser,
                                    string credentialPassword,
                                    string smtphost,
                                    int port = 25,
                                    params MailAttachment[] attachments)
        {
            try
            {
                MailMessage mail = new MailMessage
                {
                    Body = body,
                    IsBodyHtml = true
                };
                mail.To.Add(new MailAddress(to));
                mail.From = new MailAddress(fromAddress, fromDisplay, Encoding.UTF8);
                mail.Subject = subject;
                mail.SubjectEncoding = Encoding.UTF8;
                mail.Priority = MailPriority.Normal;
                foreach (MailAttachment ma in attachments)
                {
                    mail.Attachments.Add(ma.File);
                }
                SmtpClient smtp = new SmtpClient
                {
                    Host = smtphost,
                    Port = port,
                    Timeout = 50000,
                    EnableSsl = false
                };

                if (credentialUser != null && credentialPassword != null)
                    smtp.Credentials = new System.Net.NetworkCredential(credentialUser, credentialPassword);

                smtp.Send(mail);
                smtp.Dispose();
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        /// <summary>
        /// Mail without authentication
        /// </summary>
        /// <param name="to"></param>
        /// <param name="body"></param>
        /// <param name="subject"></param>
        /// <param name="fromAddress"></param>
        /// <param name="fromDisplay"></param>
        /// <param name="smtphost"></param>
        /// <param name="port"></param>
        /// <param name="attachments"></param>
        public static void NewEmail(string to,
                                    string body,
                                    string subject,
                                    string fromAddress,
                                    string fromDisplay,
                                    string smtphost,
                                    int port = 25,
                                    params MailAttachment[] attachments)
        {
            NewEmail(to, body, subject, fromAddress, fromDisplay, smtphost, port, attachments);
        }

        /*
        /// <summary>
        /// NO FUNCIONA
        /// </summary>
        /// <param name="to"></param>
        /// <param name="from"></param>
        /// <param name="body"></param>
        /// <param name="credentialUser"></param>
        /// <param name="credentialPassword"></param>
        /// <param name="smtphost"></param>
        /// <param name="port"></param>
        /// <param name="attachments"></param>
        public static void newEncryptedEmail(string to,
                                             string from,
                                             string body,
                                             string credentialUser,
                                             string credentialPassword,
                                             string smtphost,
                                             int port,
                                             string[] attachments)
        {
            string _conn = "LDAP://gd-anon.allianz.de.awin:389/O=Ced,O=GD";
            string _LdapFilter = string.Format("(&(mail={0}*))", to.Split('@')[0]);

            DirectoryEntry _entry = new DirectoryEntry(_conn, null, null, System.DirectoryServices.AuthenticationTypes.None);

            DirectorySearcher _searcher = new DirectorySearcher(_entry, _LdapFilter);

            SearchResultCollection _results = _searcher.FindAll();
            X509Certificate2 certificate1 = new X509Certificate2();

            foreach (SearchResult _res in _results)
            {
                IEnumerator enumerator2;

                enumerator2 = _res.GetDirectoryEntry().Properties["usercertificate;binary"].GetEnumerator();

                while (enumerator2.MoveNext())
                {
                    object obj1 = RuntimeHelpers.GetObjectValue(enumerator2.Current);
                    certificate1.Import((byte[])obj1);
                    //Can access different Properties for example:
                    //certificate1.Subject;
                    //certificate1.SerialNumber;
                    //certificate1.Version;
                    //certificate1.
                    //certificate1.NotAfter;
                    //certificate1.Issuer;
                    //return certificate1.Export(X509ContentType.Cert);
                }
            }

            MailMessage message = new MailMessage
            {
                From = new MailAddress(from),
                Subject = "New Administrative account \"Confidential\""
            };

            if (attachments != null && attachments.Length > 0)
            {
                StringBuilder buffer = new StringBuilder();
                buffer.Append("MIME-Version: 1.0\r\n");
                buffer.Append("Content-Type: multipart/mixed; boundary=unique-boundary-1\r\n");
                buffer.Append("\r\n");
                buffer.Append("This is a multi-part message in MIME format.\r\n");
                buffer.Append("--unique-boundary-1\r\n");
                buffer.Append("Content-Type: text/html\r\n");  //could use text/html as well here if you want a html message
                buffer.Append("Content-Transfer-Encoding: 7Bit\r\n\r\n");
                buffer.Append(body);
                if (!body.EndsWith("\r\n"))
                    buffer.Append("\r\n");
                buffer.Append("\r\n\r\n");

                foreach (string filename in attachments)
                {
                    FileInfo fileInfo = new FileInfo(filename);
                    buffer.Append("--unique-boundary-1\r\n");
                    buffer.Append("Content-Type: application/octet-stream; file=" + fileInfo.Name + "\r\n");
                    buffer.Append("Content-Transfer-Encoding: base64\r\n");
                    buffer.Append("Content-Disposition: attachment; filename=" + fileInfo.Name + "\r\n");
                    buffer.Append("\r\n");
                    byte[] binaryData = File.ReadAllBytes(filename);

                    string base64Value = Convert.ToBase64String(binaryData, 0, binaryData.Length);
                    int position = 0;
                    while (position < base64Value.Length)
                    {
                        int chunkSize = 100;
                        if (base64Value.Length - (position + chunkSize) < 0)
                            chunkSize = base64Value.Length - position;
                        buffer.Append(base64Value.Substring(position, chunkSize));
                        buffer.Append("\r\n");
                        position += chunkSize;
                    }
                    buffer.Append("\r\n");
                }

                body = buffer.ToString();
            }
            else
            {
                body = "Content-Type: text/plain\r\nContent-Transfer-Encoding: 7Bit\r\n\r\n" + body;
            }

            byte[] messageData = Encoding.ASCII.GetBytes(body);
            ContentInfo content = new ContentInfo(messageData);
            EnvelopedCms envelopedCms = new EnvelopedCms(content);
            CmsRecipientCollection toCollection = new CmsRecipientCollection();

            message.To.Add(new MailAddress(to));
            X509Certificate2 certificate = null; //Need to load from store or from file the client's cert
            CmsRecipient recipient = new CmsRecipient(SubjectIdentifierType.SubjectKeyIdentifier, certificate);
            toCollection.Add(recipient);

            envelopedCms.Encrypt(toCollection);
            byte[] encryptedBytes = envelopedCms.Encode();

            //add digital signature:
            SignedCms signedCms = new SignedCms(new ContentInfo(encryptedBytes));
            //X509Certificate2 signerCertificate = null; //Need to load from store or from file the signer's cert
            CmsSigner signer = new CmsSigner(SubjectIdentifierType.SubjectKeyIdentifier, certificate1);
            signedCms.ComputeSignature(signer);
            encryptedBytes = signedCms.Encode();
            //end digital signature section

            MemoryStream stream = new MemoryStream(encryptedBytes);
            AlternateView view = new AlternateView(stream, "application/pkcs7-mime; smime-type=signed-data;name=smime.p7m");
            message.AlternateViews.Add(view);

            SmtpClient smtp = new SmtpClient
            {
                Host = smtphost,
                Port = port,
                Credentials = new System.Net.NetworkCredential(credentialUser, credentialPassword),
                Timeout = 50000,
                EnableSsl = false
            };

            smtp.Send(message);
        }

*/

        /// <summary>
        /// Encrypted email (must have C:\PsScripts\Send-EncryptedEmail.ps1)
        /// </summary>
        /// <param name="to"></param>
        /// <param name="from"></param>
        /// <param name="body"></param>
        public static void NewEncryptedMail(string to,
                                            string from,
                                            string body)
        {
            using (PowerShell powershell = PowerShell.Create())
            {
                StringBuilder sb = new StringBuilder();

                // Add a script to the PowerShell object.
                powershell.AddCommand(@"C:\PsScripts\Send-EncryptedEmail.ps1");

                //Add parameters to previous command
                powershell.AddParameter("Emailto", to);
                powershell.AddParameter("EmailFrom", from);
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
                Console.WriteLine(sb.ToString());

                //}
            }//end using
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="to"></param>
        /// <param name="body"></param>
        /// <param name="from"></param>
        /// <param name="credentialUser"></param>
        /// <param name="credentialPassword"></param>
        /// <param name="smtphost"></param>
        /// <param name="port"></param>
        /// <param name="attachments"></param>
        public static void NewSemiPrivilegedUserEmail(string to,
                                                      string body,
                                                      string from,
                                                      string smtphost,
                                                      string credentialUser = "x",
                                                      string credentialPassword = "y",
                                                      int port = 25,
                                                      params MailAttachment[] attachments)
        {
            // Get image from Resources
            using (Bitmap bmp = new Bitmap(EguibarIT.Housekeeping.Resources.Resource1.Picture1))
            {
                // Create empty stream
                using (MemoryStream _stream = new MemoryStream())
                {
                    // Save the Bitmap to the stream
                    bmp.Save(_stream, System.Drawing.Imaging.ImageFormat.Bmp);
                    //Set the string to the start
                    _stream.Position = 0;

                    //Build the attachment
                    MailAttachment attach = new MailAttachment(_stream, "Picture1.jpg");

                    //Send the email
                    Email.NewEmail(to,
                                   body,
                                   "New Administrative account based on the AD Delegation Model",
                                   from,
                                   "Delegation Model toolset",
                                   credentialUser,
                                   credentialPassword,
                                   smtphost,
                                   25,
                                   attach);
                }
            }
        }
    }

    /// <summary>
    ///
    /// </summary>
    public class MailAttachment
    {
        #region Fields

        private readonly MemoryStream stream;
        private readonly string filename;
        private readonly string mediaType;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets the data stream for this attachment
        /// </summary>
        public Stream Data { get { return stream; } }

        /// <summary>
        /// Gets the original filename for this attachment
        /// </summary>
        public string Filename { get { return filename; } }

        /// <summary>
        /// Gets the attachment type: Bytes or String
        /// </summary>
        public string MediaType { get { return mediaType; } }

        /// <summary>
        /// Gets the file for this attachment (as a new attachment)
        /// </summary>
        public Attachment File { get { return new Attachment(Data, Filename, MediaType); } }

        /// <summary>
        ///
        /// </summary>
        public Attachment MemoryStream { get { return new Attachment(Data, MediaType); } }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Construct a mail attachment form a byte array
        /// </summary>
        /// <param name="data">Bytes to attach as a file</param>
        /// <param name="filename">Logical filename for attachment</param>
        public MailAttachment(byte[] data, string filename)
        {
            this.stream = new MemoryStream(data);
            this.filename = filename;
            this.mediaType = MediaTypeNames.Application.Octet;
        }

        /// <summary>
        /// Construct a mail attachment from a string
        /// </summary>
        /// <param name="data">String to attach as a file</param>
        /// <param name="filename">Logical filename for attachment</param>
        public MailAttachment(string data, string filename)
        {
            this.stream = new MemoryStream(System.Text.Encoding.ASCII.GetBytes(data));
            this.filename = filename;
            this.mediaType = MediaTypeNames.Text.Html;
        }

        /// <summary>
        /// Construct a mail attachment from a Memory String
        /// </summary>
        /// <param name="data">Memory stream containing the image</param>
        /// <param name="filename">Logical filename for attachment</param>
        public MailAttachment(MemoryStream data, string filename)
        {
            this.stream = data;
            this.filename = filename;
            this.mediaType = MediaTypeNames.Image.Jpeg;
        }

        #endregion Constructors
    }

    /// <summary>
    ///
    /// </summary>
    public static class HTML
    {
        #region Email body for new Semi-Privileged account

        /// <summary>
        ///
        /// </summary>
        public const string _newID = @"<!DOCTYPE html>
<html>
<head>
    <style>
/* Style Definitions */
html, body {
    font-size:14px;
    font-family:`'Roboto`',`'sans-serif`';
    color:#444;
    }
    p, li {
    margin-right:0cm;
    margin-left:0cm;
    font-size:14px;
    font-family:`'Roboto`',`'sans-serif`';
    color:#444;}
h1, h2, h3, h4 {
    font-family: 'Exo 2',sansserfi;
    color:#4678b4;
}
    </style>
</head>
<body>
    <div align = center >
        <table border=0 cellspacing=0 cellpadding=0 width=600 style='width:450.0pt'>
            <tr>
                <td>
                    <img width = 600 height=200 src=Picture1.jpg>
                </td>
            </tr>
        </table>
    <br>
    <span>
        <table border = 0 cellspacing=0 cellpadding=0 width=600 style='width:450.0pt'>
            <tr>
                <td style = 'padding:0cm 0cm 0cm 0cm' >
                    <h2><b> Announcement:</b> Operational Change on Active Directory Semi-Privileged Access Account.</h2>
                    <p>
                        As part of our continued improvement plans in our Active Directory #@DomainName@# domain, a new <i>'Delegation Model'</i>
                        is been implemented.This model will enforce a set of security guidelines that have been authorized by
                        the <b>ISO Security Team</b>, approved by our<b> Change and Release Control Committee</b> and
                        will be implemented and maintained by the <b>Infrastructure Management Teams</b>.
                    </p>
                    <p>
                        The main objective of these changes is to implement a strict segregation of duties model. In
                        Active Directory, this means that anyone who needs to manage Active Directory objects such as
                        Create/Change/Delete users, groups, computers, etc., will require a separate account, with
                        the corresponding delegated rights.These accounts are independent and not associated with your
                        standard daily usage  domain account. Below there is brief description of privileged accounts
                        in the new model.
                    </p>
                    <table border = 1 >
                        <tr border=1 style='background:black;color:silver'>
                            <th> Account </th>
                            <th> Description </th>
                        </tr>
                        <tr>
                            <td style = 'background:#ffcccc' >
                                <br>
                                SamAccountName_T0 &nbsp; &nbsp; &nbsp;
                                <br>
                            </td>
                            <td>
                                Reserved for specific restricted operational task.
                                Mainly infrastructure related.
                                Also known as 'Administration area' OR Tier0.<br>
                            </td>
                        </tr>
                        <tr>
                            <td style = 'background:#ccffcc' >
                                <br>
                                SamAccountName_T1 &nbsp; &nbsp; &nbsp;
                                <br>
                            </td>
                            <td>
                                Reserved for Servers and/or Services administration.
                                Also known as 'Servers area' OR Tier1.<br>
                            </td>
                        </tr>
                        <tr>
                            <td style = 'background:#adc6e5' >
                                <br>
                                SamAccountName_T2 &nbsp; &nbsp; &nbsp;
                                <br>
                            </td>
                            <td>
                                Reserved for standard User/Group/PC administration.
                                Also known as 'Sites area' OR Tier2.<br>
                            </td>
                        </tr>
                    </table>
                    <br>
                    <p>
                        One of the main changes to Active Directory is to implement a strict separation of permissions
                        and rights.This means that anyone who needs to manage Active Directory objects (Create/Change/Delete
                        users, groups, computers, etc.) does needs a separate account having the corresponding rights.
                    </p>
                    <p>
                    Based on your current identified role, a new administrative account has been automatically generated.
                    This account has been generated based on your current UserID (also known as SamAccountName).
                    </p>
                    <br>
                        <TABLE>
                            <TR>
                                <TD style = 'background:black;color:silver' >
                                    <br>
                                    Your new Semi-Privileged UserID is: &nbsp; &nbsp;
                                    <br><br>
                                </TD>
                                <TD style = 'background:silver' >
                                    <span style='font-size:11.0pt;font-family:'Verdana','sans-serif';color:darkblue'>
                                        &nbsp;#@UserID@#&nbsp;
                                    </span>
                                </TD>
                            </TR>
                        </TABLE>
                    <p> You will receive your password in a separate communication. </p>
                    <p>
                        As these Administrative Accounts are considered<b>'Semi-Privileged Accounts'</b> the only authorized team to
                        manage (create, reset, and remove) these accounts are the<b> Infrastructure Management Teams</b>. In the
                        event of requiring any of the previously mentioned services, please open a service request ticket to the
                        corresponding team.
                    </p>
                    <p>
                        For additional details you can get more information in our site at
                        <a href='http://www.DelegationModel.eu'> EguibarIT Delegation Model. </a>
                    </p>
                    <p> We appreciate your cooperation and collaboration in helping securing our environment. </p>
                    <p> Sincerely, </p>
                    <p> Eguibar Information Technology S.L. </p>
                </td>
            </tr>
        </table>
    </span>
</div>
<br>
<hr>
<div align = center ><font size= '2' color= '696969' face= 'arial' > This e-mail has been automatically generated by AD Delegation Model toolset</font></div>
</body>
</html>";

        #endregion Email body for new Semi-Privileged account

        #region Email body for password of Semi-Privileged account

        /// <summary>
        ///
        /// </summary>
        public const string _newPWD = @"<!DOCTYPE html>
<html>
<head>
    <style>
/* Style Definitions */
html, body {
    font-size:14px;
    font-family:`'Roboto`',`'sans-serif`';
    color:#444;
}
    p, li {
    margin-right:0cm;
    margin-left:0cm;
    font-size:14px;
    font-family:`'Roboto`',`'sans-serif`';
color:#444;}
h1, h2, h3, h4 {
    font-family: 'Exo 2',sansserfi;
    color:#4678b4;
}
    </style>
</head>
<body>
    <div align = center >
        <table border=0 cellspacing=0 cellpadding=0 width=600 style='width:450.0pt'>
            <tr>
                <td>
                    <img width = 600 height=200 src=Picture1.jpg />
                </td>
            </tr>
        </table>
    <br>
    <span>
        <table border = 0 cellspacing=0 cellpadding=0 width=600 style='width:450.0pt'>
            <tr>
                <td style = 'padding:0cm 0cm 0cm 0cm' >
                    <h2><b> Announcement: Active Directory Semi-Privileged Access Account PASSWORD.</h2>
                    <p>Following is the automatically generated password for your AD Privileged account in the domain
                    #@DomainName@#. </p>
                    <p>This password must be changed at the first logon.</p>
                        <TABLE>
                            <TR>
                                <TD style = 'background:black' >
                                    <br>
                                    <span style= 'color:silver' >
                                        Current Password: &nbsp; &nbsp;
                                    </span>
                                    <br><br>
                                </TD>
                                <TD style = 'background:silver' >
                                    <span style='color:darkblue'>
                                        &nbsp;#@unsecurePassword@#&nbsp;
                                    </span>
                                </TD>
                            </TR>
                        </TABLE>
                    <p>The password policy for these kind of accounts implement a new set of security measures.</p>
                    <ul>
                        <li>At least 14 characters long</li>
                        <li>Contain<b> UPERCASE</b> characters</li>
                        <li>Contain<b> lowercase</b> characters</li>
                        <li>Contain<b> numerical</b> characters</li>
                    </ul>
                    <br>
                    <p>
                        For additional details you can get more information in our site at
                        <a href='http://www.DelegationModel.eu'>EguibarIT Delegation Model.</a>
                    </p>
                    <p>We appreciate your cooperation and collaboration in helping securing our environment.</p>
                    <p>Sincerely,</p>
                    <p>Eguibar Information Technology S.L.</p>
                </td>
            </tr>
        </table>
    </span>
</div>
<br>
<hr>
<div align = center ><font size= '2' color= '696969' face= 'arial' > This e-mail has been automatically generated by AD Delegation Model toolset</font></div>
</body>
</html>";

        #endregion Email body for password of Semi-Privileged account

        #region Semi-Privileged Users Report

        /// <summary>
        ///
        /// </summary>
        public const string _SemiPrivilegedUsersReport = @"<!DOCTYPE html>
<html>
<head>
                 <meta charset=""UTF-8"">
                 <title>Active Directory Semi-Privileged Users Report</title>
                 <style type='text/css'>

body {
  font-family: Arial, Helvetica, sans-serif;
  font-size: 10px;
  color: #333333;
}

h1 {
  color: #3692AF;
  font-size: 18px;
  font-weight: 100;
  padding: 10px 0 2px 0;
  line-height: 12px;
}

table {
  margin: 0 0 10px 0;
  width: 90%;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.2);
  display: table;
}

th {
  height: 30px;
  border-bottom: 2px solid black;
  font-weight: 900;
  color: #ffffff;
}

thead.Tier0 {
  background: #ea6153;
}

thead.Tier1 {
  background: #27ae60;
}

thead.Tier2 {
  background: #2980b9;
}

tbody tr:nth-child(odd) {
  background: #f6f6f6;
}

tbody tr:nth-child(even) {
  background: #e9e9e9;
}

tbody tr:hoover {
  background: black;
  color: white;
}

tfoot th {
  border-bottom: 2px solid black;
}

caption {
  padding: 6px;
  font-size: 24px;
  text-align: center;
  letter-spacing: 4px;
  background: gray;
  color: white;
  font-style: bold;
}

.footer {
  position: fixed;
  left: 0;
  bottom: 0;
  width: 100%;
  background-color: #254061;
  color: white;
  text-align: center;
}

</style>
</head>
<body>
<h1>Active Directory Semi-Privileged Users Report</h1>";

        #endregion Semi-Privileged Users Report

        #region Semi-Privileged Users Report - TableHeaders

        /// <summary>
        ///
        /// </summary>
        public const string _SemiPrivilegedUsersReportTableHeaders = @"<tr>
    <th scope='col'>SamAccountName</th>
    <th scope='col'>Display Name</th>
    <th scope='col'>Description</th>
    <th scope='col'>lastLogon</th>
    <th scope='col'>pwdLastSet</th>
    <th scope='col'>pwdNeverExpires</th>
    <th scope='col'>Enabled</th>
    <th scope='col'>accountExpires</th>
    <th scope='col'>MemberOf</th>
  </tr>
  </thead>
  <tbody>
";

        #endregion Semi-Privileged Users Report - TableHeaders
    }
}