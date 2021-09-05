using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;

namespace EguibarIT.Housekeeping.AdHelper
{

    /// <summary>
    /// Class representing an generic AD object
    /// This class is used to inherit
    /// </summary>
    public class AD_Object
    {


        private DirectoryEntry DirEntry
        {
            get
            {
                if (_directoryEntry == null)
                {
                    _directoryEntry = new DirectoryEntry(LDAPPath);
                }
                return _directoryEntry;
            }
        }
        private DirectoryEntry _directoryEntry = null;

        private DirectorySearcher DirSearch
        {
            get
            {
                if (_directorySearcher == null)
                {
                    _directorySearcher = new DirectorySearcher(_directoryEntry);
                }
                return _directorySearcher;
            }
        }
        private DirectorySearcher _directorySearcher = null;

        private String LDAPPath
        {
            get
            {
                return String.Format("LDAP://{0}", AdDomain.GetdefaultNamingContext());
            }
        }


        #region Properties

        /// Global Identifier byte value returned from Active Directory and can be
        /// used to convert the identifier into different formats
        #region Guid

        private byte[] _objectGuid;

        /// <summary>
        /// The bytes that make up the identifier.
        /// </summary>
        /// <returns>A byte array</returns>
        ///<remarks>This is the format that is stored in the Active Directory repository.</remarks>
        public byte[] ObjectGuid
        {
            get
            {
                return this._objectGuid;
            }
        }

        /*
        /// <summary>
        /// The guid format of the identifier
        /// </summary>
        /// <returns>The guid format of the identifier</returns>
        public Guid Guid
        {
            get
            {
                return new Guid(this._objectGuid);
            }
        }

        /*
        /// <summary>
        /// The split octet string format of the identifier
        /// </summary>
        /// <returns>The split octet string format of the identifier</returns>
        public string GuidSplitOctetString
        {
            get
            {
                int iterator;
                StringBuilder builder;
                byte[] values = this._objectGuid;

                builder = new StringBuilder((values.GetUpperBound(0) + 1) * 2);
                for (iterator = 0; iterator <= values.GetUpperBound(0); iterator++)
                    builder.Append(@"\" + values[iterator].ToString("x2"));

                return builder.ToString();
            }
        }
        */

        /// <summary>
        /// The Guid represented as a 36 character Hex String (UUID 4)
        /// (of the form XXXXXXXX-XXXX-4XXX-YXXX-XXXXXXXXXXXX
        /// where X is any hexadecimal digit and Y is one of 8, 9, A, or B)
        /// </summary>
        public String ObjectGuidAsString
        {
            get
            {
                _objectGuidAsString = new System.Guid(ObjectGuid).ToString().ToUpper();
                return _objectGuidAsString;
            }
        }
        private String _objectGuidAsString;

        #endregion

        /// sid byte value returned from Active Directory and can be
        /// used to convert the sid into different formats.
        #region SID

        private byte[] _sid;

        /// <summary>
        /// The bytes that make up the sid.
        /// </summary>
        /// <returns>A byte array</returns>
        /// <remarks>This is the format that is stored in the Active Directory repository.</remarks>
        public byte[] SID
        {
            get
            {
                return this._sid;
            }
        }

        /*
        /// <summary>
        /// The split octet string format of the sid
        /// </summary>
        /// <returns>The split octet string format of the sid</returns>
        public string SidSplitOctetString
        {
            get
            {
                int iterator;
                StringBuilder builder;
                byte[] values = this._sid;

                builder = new StringBuilder((values.GetUpperBound(0) + 1) * 2);
                for (iterator = 0; iterator <= values.GetUpperBound(0); iterator++)
                    builder.Append(@"\" + values[iterator].ToString("x2"));

                return builder.ToString();
            }
        }
        */
        #endregion

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor method
        /// </summary>
        public AD_Object() { }



        /// <summary>
        /// Constructor method with parameters
        /// </summary>
        /// <param name="de">DirectoryEntry for the object to be loaded</param>
        public AD_Object(DirectoryEntry de)
        {
            this._objectGuid = TryGetResult<byte[]>(de, AdProperties.OBJECTGUID);
            this._sid = TryGetResult<byte[]>(de, AdProperties.OBJECTSID);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Method to convert date to AD format
        /// </summary>
        /// <param name="date">date to convert</param>
        /// <returns>string containing AD formatted date</returns>
        public static string ToADDateString(DateTime date)
        {

            string year = date.Year.ToString();
            int month = date.Month;
            int day = date.Day;
            int hr = date.Hour;
            string sb = string.Empty;
            sb += year;

            if (month < 10)
            {
                sb += "0";
            }

            sb += month.ToString();
            //bv

            if (day < 10)
            {
                sb += "0";
            }

            sb += day.ToString();

            if (hr < 10)
            {
                sb += "0";
            }

            sb += hr.ToString();
            sb += "0000.0Z";

            return sb.ToString();

        }//end ToADDateString



        /// <summary>
        /// Get the property value including its data type
        /// https://github.com/westerndigitalcorporation/DirectoryLib/blob/master/DirectoryLib/Context.cs
        /// </summary>
        /// <typeparam name="T">Type of value (int, string byte[], ...)</typeparam>
        /// <param name="de">Directory Entry object</param>
        /// <param name="key">Name of the property to retrive its value</param>
        /// <returns>AD stored value of the given property</returns>
        public T TryGetResult<T>(DirectoryEntry de, string key)
        {
            var valueCollection = de.Properties[key];

            if (valueCollection.Count > 0)
                return (T)valueCollection[0];
            else
                return default(T);
        }



        /// <summary>
        /// Override - Get the property value including its data type
        /// https://github.com/westerndigitalcorporation/DirectoryLib/blob/master/DirectoryLib/Context.cs
        /// </summary>
        /// <typeparam name="T">Type of value (int, string byte[], ...)</typeparam>
        /// <param name="de">Directory Entry object</param>
        /// <param name="key">Name of the property to retrive its value</param>
        /// <returns>AD stored value of the given property</returns>
        public T TryGetResult<T>(SearchResult de, string key)
        {
            var valueCollection = de.Properties[key];

            if (valueCollection.Count > 0)
                return (T)valueCollection[0];
            else
                return default(T);
        }



        /// <summary>
        /// Get the a list of property values including its data type
        /// https://github.com/westerndigitalcorporation/DirectoryLib/blob/master/DirectoryLib/Context.cs
        /// </summary>
        /// <typeparam name="T">Type of value (int, string byte[], ...)</typeparam>
        /// <param name="de">Directory Entry object</param>
        /// <param name="key">Name of the property to retrive its value</param>
        /// <returns>AD stored value of the given property</returns>
        public List<T> TryGetResultList<T>(DirectoryEntry de, string key)
        {
            var list = new List<T>();
            var valueCollection = de.Properties[key];

            if (valueCollection.Count > 0)
            {
                foreach (T val in valueCollection)
                {
                    list.Add(val);
                }
            }

            return list;
        }



        /// <summary>
        /// Get the user by providing DirectoryEntry
        /// </summary>
        /// <param name="de">Directory Entry</param>
        /// <returns></returns>
        //public static AD_Object GetObject(DirectoryEntry de)
        //{
        //    return new AD_Object(de);
        //}



        /// <summary>
        /// Get AD Object by Guid
        /// </summary>
        /// <param name="guid">Represents the GUID for the object we are searching for</param>
        public AD_Object GetObjectByGuid(Guid guid)
        {

            var byteArray = guid.ToByteArray();

            // to do the query, we have to format the byte array by prepending each byte with a '\'
            var hex = new StringBuilder(byteArray.Length * 3);

            foreach (byte b in byteArray)
            {
                hex.AppendFormat(@"\{0:X2}", b);
            }

            try
            {

                DirectoryEntry de = new DirectoryEntry(LDAPPath);
                DirectorySearcher ds = new DirectorySearcher(de);

                ds.Filter = string.Format(@"(&(ObjectCategory=user)(objectGuid={0}))", hex.ToString());

                SearchResult results = ds.FindOne();

                if (results == null)
                {
                    return null;
                }
                else
                {
                    //DirectoryEntry user = new DirectoryEntry(results.Path);
                    //return AD_Object(user);

                    this._objectGuid = TryGetResult<byte[]>(results, AdProperties.OBJECTGUID);
                    this._sid = TryGetResult<byte[]>(results, AdProperties.OBJECTSID);

                    return this;
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }
            finally
            {
                if(_directoryEntry != null)
                    _directoryEntry.Dispose();
                if (_directorySearcher != null)
                    _directorySearcher.Dispose();
            }
        }



        /// <summary>
        /// Get user by GUID
        /// </summary>
        /// <param name="guid">Array of 16 bytes</param>
        public AD_Object GetObjectByGuid(byte[] guid)
        {
            if (guid == null || guid.Length != 16)
            {
                throw new ArgumentException("GUID must consist of 16 bytes.");
            }

            Guid guidObj = new Guid(guid);
            return GetObjectByGuid(guidObj);
        }



        /// <summary>
        /// Get user by GUID
        /// </summary>
        /// <param name="guidString">String (hex) representation of guid (e.g. F47AC10B-58CC-4372-A567-0E02B2C3D479)</param>
        public AD_Object GetObjectByGuid(string guidString)
        {
            return GetObjectByGuid(new Guid(guidString));
        }



        #endregion

    }//end class







    /// <summary>
    /// Class representing an AD Group object
    /// </summary>
    public class AD_Group : AD_Object
    {

        #region Properties

        /// <summary>
        /// The display name for an object.
        /// </summary>
        public String DisplayName
        {
            get { return _displayname; }
            set { _displayname = value; }
        }
        private String _displayname;

        /// <summary>
        /// The email address for this contact.
        /// </summary>
        public String EmailAddress
        {
            get { return _emailAddress; }
            set { _emailAddress = value; }
        }
        private String _emailAddress;

        /// <summary>
        /// The logon name used to support clients and servers running earlier versions of the operating system.
        /// This attribute must be 20 characters or less to support earlier clients.
        /// </summary>
        public String SamAccountName
        {
            get { return _samaccountname; }
            set { _samaccountname = value; }
        }
        private String _samaccountname;

        #endregion

        #region Constructor

        /// <summary>
        /// 
        /// </summary>
        public AD_Group() { }

        #endregion

        #region Methods
        #endregion

    }//end class



    /// <summary>
    /// The class creates a wrapper for the UserAccountControl value returned from Active Directory.
    /// </summary>
    public class ADUserAccountControl
    {

        /*
        // user account control
        object value = getProperty(user, "userAccountControl");
        if (!value == null)
        {
            ADUserAccountControl wraper = new ADUserAccountControl(value);
            Console.WriteLine("Is the account disabled?:" + wraper.accountDisabled.ToString);
            Console.WriteLine("Is this a normal account?:" wraper.isNormalAccount.ToString);
            Console.WriteLine("The password doesn't expire?:" wraper.passwordNoExpire.ToString);
        }
        */


        private UserAccountControl _userAccountControlFlags;

        /// <summary>
        /// Returns the userAccountControl flags value
        /// </summary>
        /// <returns>The userAccountControl flags value</returns>
        /// <remarks>This is the value that is stored in the Active Directory repository.</remarks>
        public UserAccountControl userAccountControlFlags
        {
            get
            {
                return this._userAccountControlFlags;
            }
        }

        /// <summary>
        /// Creates a new wrapper class for the userAccontControl value
        /// </summary>
        /// <param name="userAccountControlValue">The value from active directory</param>
        public ADUserAccountControl(UserAccountControl userAccountControlValue)
        {
            this._userAccountControlFlags = userAccountControlValue;
        }

        /// <summary>
        /// Whether or not the logon script will be run.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool scriptRunOnLogin
        {
            get
            {
                return this.isFlagSet(UserAccountControl.SCRIPT);
            }
            set
            {
                this.updateFlag(UserAccountControl.SCRIPT, value);
            }
        }

        /// <summary>
        /// Whether or not the user account is disabled.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool accountDisabled
        {
            get
            {
                return this.isFlagSet(UserAccountControl.ACCOUNTDISABLE);
            }
            set
            {
                this.updateFlag(UserAccountControl.ACCOUNTDISABLE, value);
            }
        }

        /// <summary>
        /// Whether or not the home folder is required.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool homeDirectoryRequired
        {
            get
            {
                return this.isFlagSet(UserAccountControl.HOMEDIR_REQUIRED);
            }
            set
            {
                this.updateFlag(UserAccountControl.HOMEDIR_REQUIRED, value);
            }
        }

        /// <summary>
        /// Whether or not the account is locked out.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool accountLockedOut
        {
            get
            {
                return this.isFlagSet(UserAccountControl.LOCKOUT);
            }
            set
            {
                this.updateFlag(UserAccountControl.LOCKOUT, value);
            }
        }

        /// <summary>
        /// Whether or not the password is required.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool passwordNotRequired
        {
            get
            {
                return this.isFlagSet(UserAccountControl.PASSWD_NOTREQD);
            }
            set
            {
                this.updateFlag(UserAccountControl.PASSWD_NOTREQD, value);
            }
        }

        /// <summary>
        /// Whether or not the user can change his password.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        /// <remarks>
        /// Simply setting this flag will not actually change the functionality implied.  To actually make it to where a user
        /// may or may not change his password, see 
        /// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/adsi/adsi/modifying_user_cannot_change_password_ldap_provider.asp
        /// </remarks>
        public bool passwordCantChange
        {
            get
            {
                return this.isFlagSet(UserAccountControl.PASSWD_CANT_CHANGE);
            }
            set
            {
                this.updateFlag(UserAccountControl.PASSWD_CANT_CHANGE, value);
            }
        }

        /// <summary>
        /// Whether or not the user can send an encrypted password.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool encryptedTextPasswordAllowed
        {
            get
            {
                return this.isFlagSet(UserAccountControl.ENCRYPTED_TEXT_PASSWORD_ALLOWED);
            }
            set
            {
                this.updateFlag(UserAccountControl.ENCRYPTED_TEXT_PASSWORD_ALLOWED, value);
            }
        }

        /// <summary>
        /// Whether or not this is an account for users whose primary account is in another domain. This account provides 
        /// user access to this domain, but not to any domain that trusts this domain. This is sometimes referred to as a local user account.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool isTempDuplicateAccount
        {
            get
            {
                return this.isFlagSet(UserAccountControl.TEMP_DUPLICATE_ACCOUNT);
            }
            set
            {
                this.updateFlag(UserAccountControl.TEMP_DUPLICATE_ACCOUNT, value);
            }
        }

        /// <summary>
        /// Whether or not this is a default account type representing a typical user.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool isNormalAccount
        {
            get
            {
                return this.isFlagSet(UserAccountControl.NORMAL_ACCOUNT);
            }
            set
            {
                this.updateFlag(UserAccountControl.NORMAL_ACCOUNT, value);
            }
        }

        /// <summary>
        /// Whether or not this account will be trusted by domains that trust other domains.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool isInterDomainTrustAccount
        {
            get
            {
                return this.isFlagSet(UserAccountControl.INTERDOMAIN_TRUST_ACCOUNT);
            }
            set
            {
                this.updateFlag(UserAccountControl.INTERDOMAIN_TRUST_ACCOUNT, value);
            }
        }

        /// <summary>
        /// Whether or not this is a computer account for a computer that is running Microsoft Windows NT 4.0 Workstation, 
        /// Microsoft Windows NT 4.0 Server, Microsoft Windows 2000 Professional, or Windows 2000 Server and is a member of this domain.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool isWorkstationTrustAccount
        {
            get
            {
                return this.isFlagSet(UserAccountControl.WORKSTATION_TRUST_ACCOUNT);
            }
            set
            {
                this.updateFlag(UserAccountControl.WORKSTATION_TRUST_ACCOUNT, value);
            }
        }

        /// <summary>
        /// Whether or not this is a computer account for a domain controller that is a member of this domain.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool isServerTrustAccount
        {
            get
            {
                return this.isFlagSet(UserAccountControl.SERVER_TRUST_ACCOUNT);
            }
            set
            {
                this.updateFlag(UserAccountControl.SERVER_TRUST_ACCOUNT, value);
            }
        }

        /// <summary>
        /// Whether or not the password should ever expire on the account.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool passwordNoExpire
        {
            get
            {
                return this.isFlagSet(UserAccountControl.DONT_EXPIRE_PASSWD);
            }
            set
            {
                this.updateFlag(UserAccountControl.DONT_EXPIRE_PASSWD, value);
            }
        }

        /// <summary>
        /// Whether or not this is an MNS logon account.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool isMNSLogonAccount
        {
            get
            {
                return this.isFlagSet(UserAccountControl.MNS_LOGON_ACCOUNT);
            }
            set
            {
                this.updateFlag(UserAccountControl.MNS_LOGON_ACCOUNT, value);
            }
        }

        /// <summary>
        /// Whether or not the user should be forced to log on using a smart card.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool isSmartCardRequired
        {
            get
            {
                return this.isFlagSet(UserAccountControl.SMARTCARD_REQUIRED);
            }
            set
            {
                this.updateFlag(UserAccountControl.SMARTCARD_REQUIRED, value);
            }
        }

        /// <summary>
        /// When this flag is set, the service account (the user or computer account) under which a service runs is 
        /// trusted for Kerberos delegation. Any such service can impersonate a client requesting the service. To enable a 
        /// service for Kerberos delegation, you must set this flag on the <b>userAccountControl</b> property of the service account.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool isTrustedForDelegation
        {
            get
            {
                return this.isFlagSet(UserAccountControl.TRUSTED_FOR_DELEGATION);
            }
            set
            {
                this.updateFlag(UserAccountControl.TRUSTED_FOR_DELEGATION, value);
            }
        }

        /// <summary>
        /// When this flag is set, the security context of the user is not delegated to a service even if the service 
        /// account is set as trusted for Kerberos delegation.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool notDelegated
        {
            get
            {
                return this.isFlagSet(UserAccountControl.NOT_DELEGATED);
            }
            set
            {
                this.updateFlag(UserAccountControl.NOT_DELEGATED, value);
            }
        }

        /// <summary>
        /// (Windows 2000/Windows Server 2003) Whether or not this principal is restricted to use only Data Encryption Standard (DES) 
        /// encryption types for keys.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool useDesKeyOnly
        {
            get
            {
                return this.isFlagSet(UserAccountControl.USE_DES_KEY_ONLY);
            }
            set
            {
                this.updateFlag(UserAccountControl.USE_DES_KEY_ONLY, value);
            }
        }

        /// <summary>
        /// (Windows 2000/Windows Server 2003) Whether or not this account requires Kerberos pre-authentication for logging on.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool preAuthNotRequired
        {
            get
            {
                return this.isFlagSet(UserAccountControl.DONT_REQUIRE_PREAUTH);
            }
            set
            {
                this.updateFlag(UserAccountControl.DONT_REQUIRE_PREAUTH, value);
            }
        }

        /// <summary>
        /// (Windows 2000/Windows Server 2003) Whether or not the user's password has expired.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool passwordExpired
        {
            get
            {
                return this.isFlagSet(UserAccountControl.PASSWORD_EXPIRED);
            }
            set
            {
                this.updateFlag(UserAccountControl.PASSWORD_EXPIRED, value);
            }
        }

        /// <summary>
        /// (Windows 2000/Windows Server 2003) The account is enabled for delegation. This is a security-sensitive setting. 
        /// Accounts with this option enabled should be tightly controlled. This setting allows a service that runs under the 
        /// account to assume a client's identity and authenticate as that user to other remote servers on the network.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool isTrustedToAuthForDelegation
        {
            get
            {
                return this.isFlagSet(UserAccountControl.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION);
            }
            set
            {
                this.updateFlag(UserAccountControl.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION, value);
            }
        }


        private bool isFlagSet(UserAccountControl flag)
        {
            //return ((this._userAccountControlFlags AND flag) == flag);
            return ((this._userAccountControlFlags & flag).Equals(flag));
        }

        private void updateFlag(UserAccountControl flag, bool activate)
        {
            if (activate)
            {
                if (!this.isFlagSet(flag))
                    this._userAccountControlFlags = this._userAccountControlFlags & flag;
            }
            else if (this.isFlagSet(flag))
                this._userAccountControlFlags = this._userAccountControlFlags & ~flag;
        }
    }//end class



}//end namespace
