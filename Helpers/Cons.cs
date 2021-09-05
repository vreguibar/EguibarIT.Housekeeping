using System;
using System.ComponentModel;

namespace EguibarIT.Housekeeping.AdHelper
{
    /// <summary>
    /// Active Directory Properties constants
    /// </summary>
    internal class AdProperties
    {
        public const String ACCOUNTEXPIRES = "accountExpires";
        public const String ADMINCOUNT = "adminCount";
        public const String BADPASSWORDTIME = "badPasswordTime";
        public const String BADPWDCOUNT = "badPwdCount";
        public const String CITY = "l";
        public const String CODEPAGE = "codePage";
        public const String COMMENT = "comment";
        public const String COMPANY = "company";
        public const String COMMONNAME = "cn";
        public const String COUNTRY = "co";
        public const String COUNTRYCODE = "countryCode";
        public const String COUNTRYNAME = "c";
        public const String DEPARTMENT = "department";
        public const String DEPARTMENTNUMBER = "departmentNumber";
        public const String DESCRIPTION = "description";
        public const String DIRECTREPORTS = "directReports";
        public const String DISPLAYNAME = "displayName";
        public const String DISTINGUISHEDNAME = "distinguishedName";
        public const String DIVISION = "division";
        public const String DNSHOSTNAME = "dNSHostName";
        public const String DSCOREPROPAGATIONDATA = "dSCorePropagationData";
        public const String EMPLOYEEID = "employeeID";
        public const String EMPLOYEENUMBER = "employeeNumber";
        public const String EMPLOYEETYPE = "employeeType";
        public const String EMAILADDRESS = "mail";
        public const String EXTENSION = "ipPhone";
        public const String FAX = "facsimileTelephoneNumber";
        public const String FIRSTNAME = "givenName";
        public const String GROUPTYPE = "groupType";
        public const String HOMEMDIRECTORY = "homeDirectory";
        public const String HOMEMDRIVE = "homeDrive";
        public const String HOMEMDB = "homeMDB";
        public const String HOMEMTA = "homeMTA";
        public const String HOMEPHONE = "homePhone";
        public const String INSTANCETYPE = "instanceType";
        public const String LASTLOGOFF = "lastLogoff";
        public const String LASTLOGON = "lastLogon";
        public const String LASTLOGONTIMESTAMP = "lastLogonTimestamp";
        public const String LASTNAME = "sn";
        public const String LEGACYEXCHANGEDN = "legacyExchangeDN";
        public const String LOGONCOUNT = "logonCount";
        public const String MAILNICKNAME = "mailNickname";
        public const String MANAGER = "manager";
        public const String MDBUSEDEFAULTS = "mDBUseDefaults";
        public const String MEMBEROF = "memberOf";
        public const String MIDDLENAME = "initials";
        public const String MOBILE = "mobile";
        public const String MSEXCHHOMESERVERNAME = "msExchHomeServerName";
        public const String MSEXCHMAILBOXGUID = "msExchMailboxGuid";
        public const String MSEXCHMAILBOXSECURITYDESCRIPTOR = "msExchMailboxSecurityDescriptor";
        public const String MSEXCHPOLICIESINCLUDED = "msExchPoliciesIncluded";
        public const String MSEXCHRECIPIENTDISPLAYTYPE = "msExchRecipientDisplayType";
        public const String MSEXCHRECIPIENTTYPEDETAILS = "msExchRecipientTypeDetails";
        public const String MSEXCHUSERACCOUNTCONTROL = "msExchUserAccountControl";
        public const String MSEXCHVERSION = "msExchVersion";
        public const String MSNPALLOWDIALIN = "msNPAllowDialin";
        public const String NAME = "name";
        public const String NTSECURITYDESCRIPTOR = "nTSecurityDescriptor";
        public const String OBJECTCATEGORY = "objectCategory";
        public const String OBJECTCLASS = "objectClass";
        public const String OBJECTGUID = "objectGUID";
        public const String OBJECTSID = "objectSid";
        public const String OPERATINGSYSTEM = "operatingSystem";
        public const String OPERATINGSYSTEMVERSION = "OperatingSystemVersion";
        public const String ORGANIZATIONNAME = "o";
        public const String PAGER = "pager";
        public const String PERSONALTITLE = "personalTitle";
        public const String PHYSICALDELIVERYOFFICENAME = "physicalDeliveryOfficeName";
        public const String POSTALCODE = "postalCode";
        public const String POBOX = "postOfficeBox";
        public const String PRIMARYGROUPID = "primaryGroupID";
        public const String PROXYADDRESSES = "proxyAddresses";
        public const String PWDLASTSET = "pwdLastSet";
        public const String ROOMNUMBER = "roomNumber";
        public const String SAMACCOUNTNAME = "sAMAccountName";
        public const String SAMACCOUNTTYPE = "sAMAccountType";
        public const String SERVICEPRINCIPALNAME = "servicePrincipalName";
        public const String SHOWINADDRESSBOOK = "showInAddressBook";
        public const String STATE = "st";
        public const String STREETADDRESS = "street";
        public const String SUPPORTEDENCRYPTIONTYPES = "msDS-SupportedEncryptionTypes";
        public const String TELEPHONENUMBER = "telephoneNumber";
        public const String THUMBNAILPHOTO = "thumbnailPhoto";
        public const String TITLE = "title";
        public const String USERACCOUNTCONTROL = "userAccountControl";
        public const String USERPRINCIPALNAME = "userPrincipalName";
        public const String USNCHANGED = "uSNChanged";
        public const String USNCREATED = "uSNCreated";
        public const String WHENCHANGED = "whenChanged";
        public const String WHENCREATED = "whenCreated";
    }//end class

    /// <summary>
    /// Active Directory object Classess constants
    /// </summary>
    public enum objectClass
    {
        /// <summary>
        /// User Object Class
        /// </summary>
        user,

        /// <summary>
        /// Group Object Class
        /// </summary>
        group,

        /// <summary>
        /// Computer Object Class
        /// </summary>
        computer
    }

    /// <summary>
    /// Return Type constants
    /// </summary>
    public enum returnType
    {
        /// <summary>
        /// CommonName (cn)
        /// </summary>
        cn,

        /// <summary>
        /// distinguishedName
        /// </summary>
        distinguishedName,

        /// <summary>
        /// Email address
        /// </summary>
        email,

        /// <summary>
        /// Global Unique Identifier (Guid)
        /// </summary>
        Guid,

        /// <summary>
        /// GUID
        /// </summary>
        ObjectGUID,

        /// <summary>
        /// NT Name as Domain\user
        /// </summary>
        ntName,

        /// <summary>
        /// SamAccountName
        /// </summary>
        SamAccountName,

        /// <summary>
        /// Security Identifier (SID)
        /// </summary>
        SID,

        /// <summary>
        /// User Principal Name (upn)
        /// </summary>
        upn
    }

    /// <summary>
    /// These are the available flags for the userAccountControl value
    /// https://support.microsoft.com/en-us/help/305144/how-to-use-the-useraccountcontrol-flags-to-manipulate-user-account-pro
    /// </summary>
    [Flags]
    public enum UserAccountControl : int
    {
        /// <summary>
        /// The logon script is executed. (int 1)
        ///</summary>
        [Description("The logon script is executed")]
        SCRIPT = 0x0001,

        /// <summary>
        /// The user account is disabled. (int 2)
        ///</summary>
        [Description("The user account is disabled")]
        ACCOUNTDISABLE = 0x0002,

        /// <summary>
        /// The home directory is required. (int 8)
        ///</summary>
        HOMEDIR_REQUIRED = 0x0008,

        /// <summary>
        /// The account is currently locked out. (int 16)
        ///</summary>
        LOCKOUT = 0x0010,

        /// <summary>
        /// No password is required. (int 32)
        ///</summary>
        PASSWD_NOTREQD = 0x0020,

        /// <summary>
        /// The user cannot change the password. (int 64)
        ///</summary>
        /// <remarks>
        /// Note:  You cannot assign the permission settings of PASSWD_CANT_CHANGE by directly modifying the UserAccountControl attribute.
        /// For more information and a code example that shows how to prevent a user from changing the password, see User Cannot Change Password.
        /// </remarks>
        PASSWD_CANT_CHANGE = 0x0040,

        /// <summary>
        /// The user can send an encrypted password. (int 128)
        ///</summary>
        ENCRYPTED_TEXT_PASSWORD_ALLOWED = 0x0080,

        /// <summary>
        /// This is an account for users whose primary account is in another domain. This account provides user access to this domain, but not
        /// to any domain that trusts this domain. Also known as a local user account. (int 256)
        ///</summary>
        TEMP_DUPLICATE_ACCOUNT = 0x0100,

        /// <summary>
        /// This is a default account type that represents a typical user. (int 512)
        ///</summary>
        [Description("This is a default account type that represents a typical user.")]
        NORMAL_ACCOUNT = 0x0200,

        /// <summary>
        /// This is a permit to trust account for a system domain that trusts other domains. (int 2048)
        ///</summary>
        INTERDOMAIN_TRUST_ACCOUNT = 0x0800,

        /// <summary>
        /// This is a computer account for a computer that is a member of this domain. (int 4096)
        ///</summary>
        WORKSTATION_TRUST_ACCOUNT = 0x1000,

        /// <summary>
        /// This is a computer account for a system backup domain controller that is a member of this domain. (int 8192)
        ///</summary>
        SERVER_TRUST_ACCOUNT = 0x2000,

        /// <summary>
        /// Not used.
        ///</summary>
        Unused1 = 0x4000,

        /// <summary>
        /// Not used.
        ///</summary>
        Unused2 = 0x8000,

        /// <summary>
        /// The password for this account will never expire. (int 65536)
        ///</summary>
        DONT_EXPIRE_PASSWD = 0x10000,

        /// <summary>
        /// This is an MNS logon account. (int 131072)
        ///</summary>
        MNS_LOGON_ACCOUNT = 0x20000,

        /// <summary>
        /// The user must log on using a smart card. (int 262144)
        ///</summary>
        SMARTCARD_REQUIRED = 0x40000,

        /// <summary>
        /// The service account (user or computer account), under which a service runs, is trusted for Kerberos delegation. Any such service
        /// can impersonate a client requesting the service. (int 524288)
        ///</summary>
        TRUSTED_FOR_DELEGATION = 0x80000,

        /// <summary>
        /// The security context of the user will not be delegated to a service even if the service account is set as trusted for Kerberos delegation. (int 1048576)
        ///</summary>
        [Description("The security context of the user will not be delegated to a service even if the service account is set as trusted for Kerberos delegation.")]
        NOT_DELEGATED = 0x100000,

        /// <summary>
        /// Restrict this principal to use only Data Encryption Standard (DES) encryption types for keys. (int 2097152)
        ///</summary>
        USE_DES_KEY_ONLY = 0x200000,

        /// <summary>
        /// This account does not require Kerberos pre-authentication for logon. (int 4194304)
        ///</summary>
        DONT_REQUIRE_PREAUTH = 0x400000,

        /// <summary>
        /// The user password has expired. This flag is created by the system using data from the Pwd-Last-Set attribute and the domain policy. (int 8388608)
        ///</summary>
        PASSWORD_EXPIRED = 0x800000,

        /// <summary>
        /// The account is enabled for delegation. This is a security-sensitive setting; accounts with this option enabled should be strictly
        /// controlled. This setting enables a service running under the account to assume a client identity and authenticate as that user to
        /// other remote servers on the network.(int 16777216 )
        ///</summary>
        TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 0x1000000,

        /// <summary>
        /// (int 67108864 )
        /// </summary>
        PARTIAL_SECRETS_ACCOUNT = 0x4000000,

        /// <summary>
        /// (int 134217728 )
        /// </summary>
        USE_AES_KEYS = 0x08000000
    }

    /// <summary>
    /// The result after an authentication attempt.
    /// </summary>
    public enum AccountStatus
    {
        /// <summary>
        /// No user found matching the supplied account credentials. Typically a typo.
        /// </summary>
        UserNotFound = 0,

        /// <summary>
        /// The user supplied the correct credentials and successfully authenticated.
        /// </summary>
        Success,

        /// <summary>
        /// The user supplied the wrong password for the account.
        /// </summary>
        InvalidPassword,

        /// <summary>
        /// The user's password is expired and must be changed. Typically this expires every 90 days.
        /// </summary>
        ExpiredPassword,

        /// <summary>
        /// The "user must change password at next logon" since LastPasswordSet=null. Typically for new accounts.
        /// </summary>
        MustChangePassword,

        /// <summary>
        /// The user account is locked. Typically too many incorrect attempts.
        /// </summary>
        UserLockedOut
    }

    /// <summary>
    /// Login results
    /// </summary>
    public enum LoginResult
    {
        /// <summary>
        /// Login OK
        /// </summary>
        LOGIN_OK = 0,

        /// <summary>
        /// User does not exist
        /// </summary>
        LOGIN_USER_DOESNT_EXIST,

        /// <summary>
        /// User Inactive
        /// </summary>
        LOGIN_USER_ACCOUNT_INACTIVE
    }

    /// <summary>
    /// User account status
    /// </summary>
    internal enum UserStatus
    {
        /// <summary>
        /// User Enabled
        /// </summary>
        Enable = 544,

        /// <summary>
        /// User Disabled
        /// </summary>
        Disable = 546
    }

    /// <summary>
    /// Group Scope
    /// </summary>
    internal enum GroupScope
    {
        /// <summary>
        /// Domain Local Group
        /// </summary>
        ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP = -2147483644,

        /// <summary>
        /// Global Group
        /// </summary>
        ADS_GROUP_TYPE_GLOBAL_GROUP = -2147483646,

        /// <summary>
        /// Universal Group
        /// </summary>
        ADS_GROUP_TYPE_UNIVERSAL_GROUP = -2147483640
    }

    /// <summary>
    /// Defines the occurrences (or triggers) for the scheduled task
    /// </summary>
    public enum Ocurrences
    {
        /// <summary>
        /// One occurrence within 24 hours
        /// </summary>
        One,

        /// <summary>
        /// Two occurrence within 24 hours. One each 12 hours
        /// </summary>
        Two,

        /// <summary>
        /// Three occurrence within 24 hours. One each 8 hours
        /// </summary>
        Three,

        /// <summary>
        /// Four occurrence within 24 hours. One each 6 hours
        /// </summary>
        Four,

        /// <summary>
        /// Six occurrence within 24 hours. One each 4 hours
        /// </summary>
        Six,

        /// <summary>
        /// Eight occurrence within 24 hours. One each 3 hours
        /// </summary>
        Eight,

        /// <summary>
        /// Twelve occurrence within 24 hours. One each 2 hours
        /// </summary>
        Twelve,

        /// <summary>
        /// Twentyfour occurrence within 24 hours. One each 1 hours
        /// </summary>
        Twentyfour,

        /// <summary>
        /// Fortyeight occurrence within 24 hours. One each half hour
        /// </summary>
        Fortyeight
    }
}//end namespace