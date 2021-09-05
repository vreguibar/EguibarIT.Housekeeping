using System;
using System.Collections.Generic;
using System.DirectoryServices;

namespace EguibarIT.Housekeeping.AdHelper
{
    /// <summary>
    /// Class representing an AD Group object
    /// </summary>
    public class AD_Group : AD_Object
    {
        #region Properties

        /// <summary>
        /// The AdminCount attribute.
        /// If 0 object is not protected by the AdminSDHolder process
        /// </summary>
        public int adminCount { get; set; }

        /// <summary>
        /// The common name that is usually the same as DisplayName.
        /// Example: Darth Vader
        /// </summary>
        public String CommonName { get; set; }

        /// <summary>
        /// Contains the description to display for an object.
        /// This value is restricted as single-valued for backward compatibility in some cases.
        /// </summary>
        public String Description { get; set; }

        /// <summary>
        /// The display name for an object.
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// Same as the Distinguished Name for an object.
        /// </summary>
        public String DistinguishedName { get; set; }

        /// <summary>
        /// The groupType (Security or Distribution)
        /// </summary>
        public int groupType { get; set; }

        /// <summary>
        /// The email address for this contact.
        /// </summary>
        public String EmailAddress { get; set; }

        /// <summary>
        /// The name attribute
        /// </summary>
        public String name { get; set; }

        /// <summary>
        /// GUID of object
        /// </summary>
        public override byte[] ObjectGuid { get; set; }

        /// <summary>
        /// SID of object
        /// </summary>
        public override byte[] ObjectSID { get; set; }

        /// <summary>
        /// The logon name used to support clients and servers running earlier versions of the operating system.
        /// This attribute must be 20 characters or less to support earlier clients.
        /// </summary>
        public String SamAccountName { get; set; }

        #endregion Properties

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region Constructor

        /// <summary>
        ///
        /// </summary>
        public AD_Group() : base() { }

        /// <summary>
        /// Constructor method with parameters
        /// </summary>
        /// <param name="de">DirectoryEntry for the object to be loaded</param>
        public AD_Group(DirectoryEntry de) : base()
        {
            GetGroup(de);
        }

        #endregion Constructor

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region Members

        /// <summary>
        /// Returns the AD_Group object from the given DirectoryEntry
        /// </summary>
        /// <param name="de">DirectoryEntry</param>
        /// <returns>AD_Group object</returns>
        public AD_Group GetGroup(DirectoryEntry de)
        {
            this.adminCount = TryGetResult<int>(de, AdProperties.ADMINCOUNT);
            this.CommonName = TryGetResult<string>(de, AdProperties.COMMONNAME);
            this.Description = TryGetResult<string>(de, AdProperties.DESCRIPTION);
            this.DisplayName = TryGetResult<string>(de, AdProperties.DISPLAYNAME);
            this.DistinguishedName = TryGetResult<string>(de, AdProperties.DISTINGUISHEDNAME);
            this.groupType = TryGetResult<int>(de, AdProperties.GROUPTYPE);
            this.EmailAddress = TryGetResult<string>(de, AdProperties.EMAILADDRESS);
            this.name = TryGetResult<string>(de, AdProperties.NAME);
            this.SamAccountName = TryGetResult<string>(de, AdProperties.SAMACCOUNTNAME);
            this.ObjectGuid = TryGetResult<byte[]>(de, AdProperties.OBJECTGUID);
            this.ObjectSID = TryGetResult<byte[]>(de, AdProperties.OBJECTSID);

            return this;
        }

        /// <summary>
        /// Get Group by Guid
        /// </summary>
        /// <param name="guid">Represents the GUID for the object we are searching for</param>
        /// <returns>AD_Group object</returns>
        public AD_Group GetGroupByGuid(Guid guid)
        {
            try
            {
                using (DirectorySearcher DirSearch = GetDirectorySearcher())
                {
                    DirSearch.Filter = String.Format("(&(objectClass=group)(objectGUID={0}))", GuidToHex(guid));

                    SearchResult results = DirSearch.FindOne();

                    if (results == null)
                    {
                        return null;
                    }
                    else
                    {
                        //DirectoryEntry user = new DirectoryEntry(results.Path);

                        //return GetGroup(new DirectoryEntry(results.Path));
                        return GetGroup(GetDirectoryEntry(results.Path));
                    }
                }//end using
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }
        }

        /// <summary>
        /// Take a Group and return list of member users
        /// </summary>
        /// <param name="groupName"></param>
        /// <returns></returns>
        public List<AD_UserFull> GetUserFromGroup(String groupName)
        {
            List<AD_UserFull> userlist = new List<AD_UserFull>();
            try
            {
                using (DirectorySearcher DirSearch = GetDirectorySearcher())
                {
                    DirSearch.Filter = string.Format("(&(objectClass=group)(SAMAccountName={0}", groupName);

                    SearchResult results = DirSearch.FindOne();

                    if (results != null)
                    {
                        DirectoryEntry deGroup = new DirectoryEntry(results.Path);

                        System.DirectoryServices.PropertyCollection pColl = deGroup.Properties;

                        int count = pColl["member"].Count;

                        for (int i = 0; i < count; i++)
                        {
                            string respath = results.Path;
                            string[] pathnavigate = respath.Split("CN".ToCharArray());
                            respath = pathnavigate[0];
                            string objpath = pColl["member"][i].ToString();
                            string path = respath + objpath;

                            AD_UserFull userobj = new AD_UserFull(new DirectoryEntry(path));

                            userlist.Add(userobj);
                        }
                    }
                }//end using
                return userlist;
            }
            catch (Exception ex)
            {
                ex.ToString();
                return userlist;
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="ouPath"></param>
        /// <param name="name"></param>
        public void Create(string ouPath, string name)
        {
            if (!DirectoryEntry.Exists("LDAP://CN=" + name + "," + ouPath))
            {
                try
                {
                    DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                    DirectoryEntry group = entry.Children.Add("CN=" + name, "group");
                    group.Properties["sAmAccountName"].Value = name;
                    group.CommitChanges();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message.ToString());
                }
            }
            else { Console.WriteLine(ouPath + " already exists"); }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="ouPath"></param>
        /// <param name="groupPath"></param>
        public void Delete(string ouPath, string groupPath)
        {
            if (DirectoryEntry.Exists("LDAP://" + groupPath))
            {
                try
                {
                    DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                    DirectoryEntry group = new DirectoryEntry("LDAP://" + groupPath);
                    entry.Children.Remove(group);
                    group.CommitChanges();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message.ToString());
                }
            }
            else
            {
                Console.WriteLine(ouPath + " doesn't exist");
            }
        }

        #endregion Members
    }//end class
}