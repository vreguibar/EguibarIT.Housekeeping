using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Text;

namespace EguibarIT.Housekeeping.AdHelper
{
    /// <summary>
    ///
    /// </summary>
    public class Helpers
    {
        /*
        /// <summary>
        ///
        /// </summary>
        protected DirectoryEntry DirEntry
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

        /// <summary>
        ///
        /// </summary>
        protected DirectorySearcher DirSearch
        {
            get
            {
                if (_directorySearcher == null)
                {
                    _directorySearcher = new DirectorySearcher(DirEntry);
                }
                return _directorySearcher;
            }
        }
        private DirectorySearcher _directorySearcher = null;

        /// <summary>
        ///
        /// </summary>
        protected String LDAPPath
        {
            get
            {
                return String.Format("LDAP://{0}", AdDomain.GetdefaultNamingContext());
            }
        }
        */

        #region Get/Set Properties

        /// <summary>
        /// Get the property value by providing its name from SearchResult
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="result">SearchResult</param>
        /// <param name="key">Name of propperty to get its value</param>
        /// <returns>Value of the property</returns>
        protected T TryGetResult<T>(SearchResult result, string key)
        {
            var valueCollection = result.Properties[key];
            if (valueCollection.Count > 0)
                return (T)valueCollection[0];
            else
                return default(T);
        }

        /// <summary>
        /// Get the property value by providing its name from DirectoryEntry
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="result">DirectoryEntry</param>
        /// <param name="key">Name of propperty to get its value</param>
        /// <returns>Value of the property</returns>
        protected T TryGetResult<T>(DirectoryEntry result, string key)
        {
            var valueCollection = result.Properties[key];
            if (valueCollection.Count > 0)
                return (T)valueCollection[0];
            else
                return default(T);
        }

        /// <summary>
        /// Set the property value
        /// </summary>
        /// <param name="de">DirectoryEntry</param>
        /// <param name="key">Name of propperty</param>
        /// <param name="PropertyValue">New value of the property</param>
        /// <returns></returns>
        protected void TrySetProperty(DirectoryEntry de, string key, string PropertyValue)
        {
            try
            {
                if ((PropertyValue != string.Empty) && (PropertyValue != null))
                {
                    if (de.Properties.Contains(key))
                    {
                        de.Properties[key][0] = PropertyValue;
                    }
                    else
                    {
                        de.Properties[key].Add(PropertyValue);
                    }
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        /// <summary>
        /// Get the list of property values by providing its name from SearchResult
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="result">SearchResult</param>
        /// <param name="key">Name of propperty to get its value</param>
        /// <returns>List of values from the property</returns>
        protected List<T> TryGetResultList<T>(SearchResult result, string key)
        {
            var list = new List<T>();
            var valueCollection = result.Properties[key];
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
        /// Get the list of property values by providing its name from DirectoryEntry
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="result">DirectoryEntry</param>
        /// <param name="key">Name of propperty to get its value</param>
        /// <returns>List of values from the property</returns>
        protected List<T> TryGetResultList<T>(DirectoryEntry result, string key)
        {
            var list = new List<T>();
            var valueCollection = result.Properties[key];
            if (valueCollection.Count > 0)
            {
                foreach (T val in valueCollection)
                {
                    list.Add(val);
                }
            }
            return list;
        }

        #endregion Get/Set Properties

        #region GUID conversions

        /// <summary>
        /// Overload - Convert System GUID to Octet (HEX) string.
        /// </summary>
        /// <param name="guid">System GUID</param>
        /// <returns>HEX string representing GUID, including "\"</returns>
        protected static string GuidToHex(Guid guid)
        {
            //Call GuidToHex passing the ByteArray
            return GuidToHex(guid.ToByteArray());
        }

        /// <summary>
        /// Overload - Convert String GUID to Octet (HEX) string.
        /// </summary>
        /// <param name="guid">String GUID</param>
        /// <returns>HEX string representing GUID, including "\"</returns>
        protected static string GuidToHex(string guid)
        {
            //Convert string into GUID
            System.Guid _guid = new Guid(guid);

            //Call GuidToHex passing the ByteArray
            return GuidToHex(_guid.ToByteArray());
        }

        /// <summary>
        /// Convert ByteArray GUID to Octet (HEX) string.
        /// </summary>
        /// <param name="guid">ByteArray GUID</param>
        /// <returns>HEX string representing GUID, including "\"</returns>
        protected static string GuidToHex(byte[] guid)
        {
            if (guid == null || guid.Length != 16)
            {
                throw new ArgumentException("GUID must consist of 16 bytes.");
            }

            // to do the query, we have to format the byte array by prepending each byte with a '\'
            var hex = new StringBuilder(guid.Length * 3);

            foreach (byte b in guid)
            {
                hex.AppendFormat(@"\{0:X2}", b);
            }
            return hex.ToString();
        }

        /// <summary>
        /// Convert System GUID to a readable format (ej. b75f1fdd-2294-40e4-85df-0ce0e39895be )
        /// </summary>
        /// <param name="guid">System GUID</param>
        /// <returns>String in readable format representing the GUID</returns>
        protected static string GuidToReadable(Guid guid)
        {
            /*
             https://msdn.microsoft.com/en-us/library/system.guid.parse(v=vs.110).aspx
             N =                  32 digits                              = 00000000000000000000000000000000
             D =            32 digits separated by hyphens               = 00000000-0000-0000-0000-000000000000
             B =    32 digits separated by hyphens, enclosed in braces   = {00000000-0000-0000-0000-000000000000}
             P = 32 digits separated by hyphens, enclosed in parentheses = (00000000-0000-0000-0000-000000000000)
             X = Four hexadecimal values enclosed in braces,
                  where the fourth value is a subset of eight
                  hexadecimal values that is also enclosed in braces     = {0x00000000,0x0000,0x0000,{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00}}
             */
            return guid.ToString("D");
        }

        /// <summary>
        /// Convert ByteArray GUID to a readable format (ej. b75f1fdd-2294-40e4-85df-0ce0e39895be )
        /// </summary>
        /// <param name="guid">ByteArray GUID</param>
        /// <returns>String in readable format representing the GUID</returns>
        protected static string GuidToReadable(byte[] guid)
        {
            /*
             https://msdn.microsoft.com/en-us/library/system.guid.parse(v=vs.110).aspx
             N =                  32 digits                              = 00000000000000000000000000000000
             D =            32 digits separated by hyphens               = 00000000-0000-0000-0000-000000000000
             B =    32 digits separated by hyphens, enclosed in braces   = {00000000-0000-0000-0000-000000000000}
             P = 32 digits separated by hyphens, enclosed in parentheses = (00000000-0000-0000-0000-000000000000)
             X = Four hexadecimal values enclosed in braces,
                  where the fourth value is a subset of eight
                  hexadecimal values that is also enclosed in braces     = {0x00000000,0x0000,0x0000,{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00}}
             */
            Guid newGuid = new Guid(guid);

            return newGuid.ToString("D");
        }

        #endregion GUID conversions

        #region DirectoryEntry

        /// <summary>
        /// Get the DirectoryEntry from given LdapPath (ej. "LDAP://DC=EguibarIT,DC=local")
        /// </summary>
        /// <param name="LdapPath">"LDAP://DC=EguibarIT,DC=local"</param>
        /// <returns>DirectoryEntry</returns>
        protected static DirectoryEntry GetDirectoryEntry(string LdapPath)
        {
            if (!LdapPath.StartsWith("LDAP://"))
            {
                LdapPath = string.Format("LDAP://{0}", LdapPath);
            }

            DirectoryEntry de = new DirectoryEntry(LdapPath);
            return de;
        }

        /// <summary>
        /// Overload - Get the DirectoryEntry from default LdapPath of current domain
        /// </summary>
        /// <returns>DirectoryEntry</returns>
        protected static DirectoryEntry GetDirectoryEntry()
        {
            DirectoryEntry de = new DirectoryEntry(String.Format("LDAP://{0}", AdDomain.GetdefaultNamingContext()));
            return de;
        }

        /// <summary>
        /// Get DirectorySearcher from the given DirectoryEntry
        /// </summary>
        /// <param name="de">DirectoryEntry</param>
        /// <returns>DirectorySearcher</returns>
        protected static DirectorySearcher GetDirectorySearcher(DirectoryEntry de)
        {
            DirectorySearcher ds = new DirectorySearcher(de);
            return ds;
        }

        /// <summary>
        /// Overload - Get DirectorySearcher from the default DirectoryEntry
        /// </summary>
        /// <returns>DirectorySearcher</returns>
        protected static DirectorySearcher GetDirectorySearcher()
        {
            DirectoryEntry de = GetDirectoryEntry();
            DirectorySearcher ds = new DirectorySearcher(de);
            return ds;
        }

        #endregion DirectoryEntry

        /// <summary>
        /// Verify if object exists
        /// </summary>
        /// <param name="objectPath">LDAP path to the given object</param>
        /// <returns>Bool</returns>
        public static bool Exists(string objectPath)
        {
            bool found = false;
            if (DirectoryEntry.Exists("LDAP://" + objectPath))
            {
                found = true;
            }
            return found;
        }

        /// <summary>
        /// Move an object
        /// </summary>
        /// <param name="objectLocation">Original DN of the object</param>
        /// <param name="newLocation">Destination DN where the object will be moved to.</param>
        public void Move(string objectLocation, string newLocation)
        {
            //For brevity, removed existence checks

            DirectoryEntry eLocation = new DirectoryEntry("LDAP://" + objectLocation);
            DirectoryEntry nLocation = new DirectoryEntry("LDAP://" + newLocation);
            string newName = eLocation.Name;
            eLocation.MoveTo(nLocation, newName);
            nLocation.Close();
            eLocation.Close();
        }

        /// <summary>
        /// Rename an object
        /// </summary>
        /// <param name="objectDn">Object DN to be renamed</param>
        /// <param name="newName">New Name of the object</param>
        public static void Rename(string objectDn, string newName)
        {
            DirectoryEntry child = new DirectoryEntry("LDAP://" + objectDn);
            child.Rename("CN=" + newName);
        }

        //http://www.informit.com/articles/article.aspx?p=474649&seqNum=3
        //
    }

    /// <summary>
    /// The class creates a wrapper for the UserAccountControl value returned from Active Directory.
    /// </summary>
    public class ADUserAccountControl
    {
        /// <summary>
        /// Returns the userAccountControl flags value
        /// </summary>
        /// <returns>The userAccountControl flags value</returns>
        /// <remarks>This is the value that is stored in the ENUM.</remarks>
        public UserAccountControl userAccountControlFlags { get; private set; }

        /// <summary>
        /// Creates a new wrapper class for the userAccontControl value
        /// </summary>
        /// <param name="userAccountControlValue">The value from active directory</param>
        public ADUserAccountControl(UserAccountControl userAccountControlValue)
        {
            this.userAccountControlFlags = userAccountControlValue;
        }

        /// <summary>
        ///
        /// </summary>
        public ADUserAccountControl()
        {
        }

        #region Properties

        /// <summary>
        /// Whether or not the logon script will be run.
        /// </summary>
        /// <value>Set true to activate and false to deactivate</value>
        /// <returns>Whether or not the flag is active</returns>
        public bool scriptRunOnLogin
        {
            get
            {
                return this.userAccountControlFlags.HasFlag(UserAccountControl.SCRIPT);
                //return this.isFlagSet(UserAccountControl.SCRIPT);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.ACCOUNTDISABLE);
                //return this.isFlagSet(UserAccountControl.ACCOUNTDISABLE);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.HOMEDIR_REQUIRED);
                //return this.isFlagSet(UserAccountControl.HOMEDIR_REQUIRED);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.LOCKOUT);
                //return this.isFlagSet(UserAccountControl.LOCKOUT);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.PASSWD_NOTREQD);
                //return this.isFlagSet(UserAccountControl.PASSWD_NOTREQD);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.PASSWD_CANT_CHANGE);
                //return this.isFlagSet(UserAccountControl.PASSWD_CANT_CHANGE);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.ENCRYPTED_TEXT_PASSWORD_ALLOWED);
                //return this.isFlagSet(UserAccountControl.ENCRYPTED_TEXT_PASSWORD_ALLOWED);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.TEMP_DUPLICATE_ACCOUNT);
                //return this.isFlagSet(UserAccountControl.TEMP_DUPLICATE_ACCOUNT);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.INTERDOMAIN_TRUST_ACCOUNT);
                //return this.isFlagSet(UserAccountControl.INTERDOMAIN_TRUST_ACCOUNT);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.INTERDOMAIN_TRUST_ACCOUNT);
                //return this.isFlagSet(UserAccountControl.INTERDOMAIN_TRUST_ACCOUNT);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.WORKSTATION_TRUST_ACCOUNT);
                //return this.isFlagSet(UserAccountControl.WORKSTATION_TRUST_ACCOUNT);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.SERVER_TRUST_ACCOUNT);
                //return this.isFlagSet(UserAccountControl.SERVER_TRUST_ACCOUNT);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.DONT_EXPIRE_PASSWD);
                //return this.isFlagSet(UserAccountControl.DONT_EXPIRE_PASSWD);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.MNS_LOGON_ACCOUNT);
                //return this.isFlagSet(UserAccountControl.MNS_LOGON_ACCOUNT);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.SMARTCARD_REQUIRED);
                //return this.isFlagSet(UserAccountControl.SMARTCARD_REQUIRED);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.TRUSTED_FOR_DELEGATION);
                //return this.isFlagSet(UserAccountControl.TRUSTED_FOR_DELEGATION);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.NOT_DELEGATED);
                //return this.isFlagSet(UserAccountControl.NOT_DELEGATED);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.USE_DES_KEY_ONLY);
                //return this.isFlagSet(UserAccountControl.USE_DES_KEY_ONLY);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.DONT_REQUIRE_PREAUTH);
                //return this.isFlagSet(UserAccountControl.DONT_REQUIRE_PREAUTH);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.PASSWORD_EXPIRED);
                //return this.isFlagSet(UserAccountControl.PASSWORD_EXPIRED);
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
                return this.userAccountControlFlags.HasFlag(UserAccountControl.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION);
                //return this.isFlagSet(UserAccountControl.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION);
            }
            set
            {
                this.updateFlag(UserAccountControl.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION, value);
            }
        }

        #endregion Properties

        //https://www.alanzucconi.com/2015/07/26/enum-flags-and-bitwise-operators/

        private void updateFlag(UserAccountControl flag, bool activate)
        {
            if (activate)
            {
                if (!this.userAccountControlFlags.HasFlag(flag))
                    this.userAccountControlFlags = this.userAccountControlFlags & flag;
            }
            else if (this.userAccountControlFlags.HasFlag(flag))
                this.userAccountControlFlags = this.userAccountControlFlags & ~flag;
        }

        private static UserAccountControl ToogleFlag(UserAccountControl a, UserAccountControl b)
        {
            return a ^ b;
        }
    }//end class

    /*
    /// <summary>
    /// Enum Extensions
    /// Overloaded Members for Flag Management
    /// </summary>
    public static class EnumExtensions
    {
        /// <summary>
        /// Check if Enum
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="withFlags">Flags to compare against.</param>
        private static void CheckIsEnum<T>(bool withFlags)
        {
            if (!typeof(T).IsEnum)
                throw new ArgumentException(string.Format("Type '{0}' is not an enum", typeof(T).FullName));
            if (withFlags && !Attribute.IsDefined(typeof(T), typeof(FlagsAttribute)))
                throw new ArgumentException(string.Format("Type '{0}' doesn't have the 'Flags' attribute", typeof(T).FullName));
        }

        /// <summary>
        /// Check Is the Flag Set
        /// </summary>
        /// <typeparam name="T">The type of the enum.</typeparam>
        /// <param name="value">Value of the given Type</param>
        /// <param name="flag">Flags to compare against.</param>
        /// <returns>
        ///  <c>true</c> if the specified value has flags; otherwise, <c>false</c>.
        /// </returns>
        /// (a & b) == b
        public static bool IsFlagSet<T>(this T value, T flag) where T : struct, IConvertible, IComparable, IFormattable
        {
            CheckIsEnum<T>(true);
            long lValue = Convert.ToInt64(value);
            long lFlag = Convert.ToInt64(flag);
            return (lValue & lFlag) != 0;
        }

        /// <summary>
        /// Get the Flags
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="value">Value of the given Type</param>
        /// <returns>IEnumerable object containing the Flags</returns>
        public static IEnumerable<T> GetFlags<T>(this T value) where T : struct, IConvertible, IComparable, IFormattable
        {
            CheckIsEnum<T>(true);
            foreach (T flag in Enum.GetValues(typeof(T)).Cast<T>())
            {
                if (value.IsFlagSet(flag))
                    yield return flag;
            }
        }

        /// <summary>
        /// Set the Flags
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="value">Value of the given Type</param>
        /// <param name="flags">Flags to set up.</param>
        /// <param name="on">Turn ON the flag</param>
        /// <returns></returns>
        /// a | b  OR  a & (~b)
        public static T SetFlags<T>(this T value, T flags, bool on) where T : struct, IConvertible, IComparable, IFormattable
        {
            CheckIsEnum<T>(true);
            long lValue = Convert.ToInt64(value);
            long lFlag = Convert.ToInt64(flags);
            if (on)
            {
                //Set flag
                lValue |= lFlag;
            }
            else
            {
                //Clear flag
                //a & (~b)
                lValue &= (~lFlag);
            }
            return (T)Enum.ToObject(typeof(T), lValue);
        }

        /// <summary>
        /// Overload - Set the Flags
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="value">Value of the given Type</param>
        /// <param name="flags">Flags to set up.</param>
        /// <returns></returns>
        public static T SetFlags<T>(this T value, T flags) where T : struct, IConvertible, IComparable, IFormattable
        {
            return value.SetFlags(flags, true);
        }

        /// <summary>
        /// Clear the flag
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="value">Value of the given Type</param>
        /// <param name="flags">Flags to clear up.</param>
        /// <returns></returns>
        /// a & (~b)
        public static T ClearFlags<T>(this T value, T flags) where T : struct, IConvertible, IComparable, IFormattable
        {
            return value.SetFlags(flags, false);
        }

        /// <summary>
        /// Combine flags.
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="flags">Flags to combine.</param>
        /// <returns></returns>
        public static T CombineFlags<T>(this IEnumerable<T> flags) where T : struct, IConvertible, IComparable, IFormattable
        {
            CheckIsEnum<T>(true);
            long lValue = 0;
            foreach (T flag in flags)
            {
                long lFlag = Convert.ToInt64(flag);
                lValue |= lFlag;
            }
            return (T)Enum.ToObject(typeof(T), lValue);
        }

        /*
        public static string GetDescription<T>(this T value) where T : struct, IConvertible, IComparable, IFormattable
        {
            CheckIsEnum<T>(false);
            string name = Enum.GetName(typeof(T), value);
            if (name != null)
            {
                FieldInfo field = typeof(T).GetField(name);
                if (field != null)
                {
                    DescriptionAttribute attr = Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) as DescriptionAttribute;
                    if (attr != null)
                    {
                        return attr.Description;
                    }
                }
            }
            return null;
        }

    //}////////
*/
}//end namespace