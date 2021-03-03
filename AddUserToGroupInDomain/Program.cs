using System;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace AddUserToGroupInDomain
{
    // Ref - https://msdn.microsoft.com/en-us/library/ms676310(v=vs.85).aspx
    internal class Program
    {
        // Problem Statement - Add user from domain.com to group in child.domain.com
            
        private const string ParentDomain = "domain.com";
        private const string ChildDomain = "child.domain.com";
        //private const string Container = "CN=Users,DC=domain,DC=com";
        private const string Username = @"domain\Administrator";
        private const string ChildUsername = @"domain\Administrator";
        private const string Password = "Pass99";

        private static void Main()
        {
            //const string bindString = "LDAP://domain.com/CN=TicketGroup,OU=Groups,DC=domain,DC=com";
            //const string newMember = "CN=Vikram Singh Saini,OU=Jersey,DC=domain,DC=com";
            //AddMemberToGroup(bindString, newMember);

            const string userToAdd = "vssaini";
            const string childGroup = "Child Group";
            WriteToConsole("Initiating process ...");
            Console.WriteLine();

            try
            {
                AddUserToGroup(userToAdd, childGroup);
            }
            catch (Exception e)
            {
                WriteToConsole(e.ToString(),true);
            }
            
            Console.WriteLine();
            WriteToConsole("Process completed.");
            Console.ReadKey();
        }

        /// <summary>
        /// Add new member to group in domain.
        /// </summary>
        /// <param name="bindString">A valid ADsPath for a group container</param>
        /// <param name="newMember">The distinguished name of the member to be added to the group</param>
        public static void AddMemberToGroup(string bindString, string newMember)
        {
            try
            {
                var ent = new DirectoryEntry(bindString);
                ent.Properties["member"].Add(newMember);
                ent.CommitChanges();

                Console.WriteLine("Member added to domain successfully!");
            }

            catch (Exception e)
            {
                Console.WriteLine("An error occurred.");
                Console.WriteLine("{0}", e.Message);
            }
        }

        /// <summary>
        /// Add member from one group in parent domain to another group in child domain.
        /// </summary>
        /// <param name="upn">The user from parent domain.</param>
        /// <param name="groupName">The name of the group in child domain.</param>
        private static void AddUserToGroup(string upn, string groupName)
        {
            var domainContext = GetPrincipalContext();
            WriteToConsole($"Searching for parent user {upn} in domain.com");

            //GET THE USER FROM DOMAIN domain.com
            using (var parentUser = UserPrincipal.FindByIdentity(domainContext, upn))
            {
                if (parentUser != null)
                {
                    WriteToConsole($"Parent user {parentUser.SamAccountName} was found.");
                    WriteToConsole($"Searching group {groupName} in {ChildDomain}");

                    //FIND THE GROUP IN DOMAIN child.domain.com
                    var childDomainContext = GetChildPrincipalContext();
                    var childGroupPrincipal = GroupPrincipal.FindByIdentity(childDomainContext, groupName);
                    using (childGroupPrincipal)
                    {
                        if (childGroupPrincipal != null)
                        {
                            WriteToConsole($"Child group {groupName} found in {ChildDomain}");

                            //CHECK TO MAKE SURE USER IS NOT IN THAT GROUP
                            if (!parentUser.IsMemberOf(childGroupPrincipal))
                            {
                                // Ref for server is unwilling to process request
                                // - http://stackoverflow.com/questions/13748970/server-is-unwilling-to-process-the-request-active-directory-add-user-via-c-s

                                var userDn = parentUser.DistinguishedName;
                                var userDnFullPath = $"LDAP://{ParentDomain}/{userDn}";
                                //var userSid = $"<SID={parentUser.Sid}>";
                                var childGroupDe = (DirectoryEntry)childGroupPrincipal.GetUnderlyingObject();
                                childGroupDe.Invoke("Add", userDnFullPath);
                                //groupDirectoryEntry.Properties["member"].Add(userSid);
                                childGroupDe.CommitChanges();

                                WriteToConsole($"Parent user {parentUser.SamAccountName} added to child group {groupName} successfully.");
                            }
                            else
                            {
                                WriteToConsole($"User {parentUser.SamAccountName} is already member of  group {groupName}");
                            }
                        }
                        else
                        {
                            WriteToConsole($"Child group {groupName} not found in {ChildDomain}", true);
                        }
                    }

                }
            }
        }

        /// <summary>
        /// Gets the parent principal context
        /// </summary>
        /// <returns>Returns the PrincipalContext object</returns>
        public static PrincipalContext GetPrincipalContext()
        {
            WriteToConsole("Creating parent domain context");
            var principalContext = new PrincipalContext(ContextType.Domain, ParentDomain, Username, Password);
            return principalContext;
        }

        /// <summary>
        /// Gets the child principal context
        /// </summary>
        /// <returns>Returns the PrincipalContext object</returns>
        public static PrincipalContext GetChildPrincipalContext()
        {
            WriteToConsole("Creating child domain context");
            var principalContext = new PrincipalContext(ContextType.Domain, ChildDomain, ChildUsername, Password);
            return principalContext;
        }
        
        /// <summary>
        /// Write message to conosole.
        /// </summary>
        /// <param name="msg">The message that need to be written to console.</param>
        /// <param name="isError">Whether to show the message in red color.</param>
        private static void WriteToConsole(string msg, bool isError = false)
        {
            if (isError)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(msg);
                Console.ResetColor();
            }
            else
            {
                Console.WriteLine(msg);
            }
        }
    }
}
