using System.DirectoryServices.AccountManagement;

namespace EguibarIT.Housekeeping.AdHelper
{
    //https://www.c-sharpcorner.com/uploadfile/dhananjaycoder/all-operations-on-active-directory-ad-using-C-Sharp/

    /// <summary>
    ///
    /// </summary>
    public class ADManager
    {
        private PrincipalContext context;

        /// <summary>
        ///
        /// </summary>
        public ADManager()
        {
            context = new PrincipalContext(ContextType.Domain);
        }

        /// <summary>
        /// Set the domain context
        /// </summary>
        /// <param name="domain"></param>
        public ADManager(string domain)
        {
            context = new PrincipalContext(ContextType.Domain, domain);
        }

        /// <summary>
        /// Set the domain context and container
        /// </summary>
        /// <param name="domain"></param>
        /// <param name="container"></param>
        public ADManager(string domain, string container)
        {
            context = new PrincipalContext(ContextType.Domain, domain, container);
        }

        /// <summary>
        /// Set the domain context providing user and password
        /// </summary>
        /// <param name="domain">Domain name.</param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        public ADManager(string domain, string username, string password)
        {
            context = new PrincipalContext(ContextType.Domain, username, password);
        }

        /// <summary>
        /// Add a user to a group.
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="groupName"></param>
        /// <returns></returns>
        public bool AddUserToGroup(string userName, string groupName)
        {
            bool done = false;

            GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);
            if (group == null)
            {
                group = new GroupPrincipal(context, groupName);
            }
            UserPrincipal user = UserPrincipal.FindByIdentity(context, userName);
            if (user != null & group != null)
            {
                group.Members.Add(user);
                group.Save();
                done = (user.IsMemberOf(group));
            }
            return done;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="groupName"></param>
        /// <returns></returns>
        public bool RemoveUserFromGroup(string userName, string groupName)
        {
            bool done = false;
            UserPrincipal user = UserPrincipal.FindByIdentity(context, userName);
            GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);
            if (user != null & group != null)
            {
                group.Members.Remove(user);
                group.Save();
                done = !(user.IsMemberOf(group));
            }
            return done;
        }
    }//end class
}//end namespace