using NLog;
using SelectHealth.Ops.AWD.Core;
using SelectHealth.Ops.AWD.Entities;
using SelectHealth.Ops.AWD.Logger;
using System.Net;
using NPOI;
using NPOI.SS.UserModel;

namespace AWD_Create_Helper
{
    internal class Program
    {
        private static string awd_url;
        private static string auth_username;
        private static string template_username;
        private static Logger logger;
        private static SH_NLogger wrapped_logger;
        private static RequestService request;
        private static AuthenticationService auth;
        private static AccountService account;
        private static HttpClient client;
        private static Credentials credentials = null;

        static async Task Main(string[] args)
        {
            try
            {
                // set up the environment
                await Setup_Env();

                var users = Read_And_Get_Users_From_Spreadsheet("TestUsers.xlsx");
                //var example_user = Return_Example_Template_User();
                foreach (var user in users)
                {
                    var success_in_creating_user = await Create_User(user);
                    if (success_in_creating_user)
                        await Set_Workspace_to_Admin(user.Username);
                }
            }
            finally
            {
                if (credentials != null)
                    await auth.SignOut(credentials);
                Console.WriteLine("Signed out");
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadLine();
        }

        /// <summary>
        /// Set the default workspace to Administrator. The user must be on cyberark for this method to work. 
        /// </summary>
        /// <param name="username"></param>
        private static async Task Set_Workspace_to_Admin(string username)
        {
            var password = await auth.GetCyberarkPassword(username, "DEV01");
            Credentials credentials = null;
            try
            {
                credentials = await auth.SignIn(username, password);
                var request = new HttpRequestMessage(HttpMethod.Put, $"{awd_url}services/v1/user/workspace?name=PROCVIEW&_=1733331099811");
                request.Headers.Add("csrf_token", credentials.csrf_token);
                request.Headers.Add("Cookie", "JSESSIONID=" + credentials.jsession_cookie);

                var response = await client.SendAsync(request);
                if (response.StatusCode == HttpStatusCode.NoContent)
                    Console.WriteLine($"Successfully set workspace for {username} to Administrator");
                else
                    Console.WriteLine("Failed to set workspace to Administrator");
            }
            finally
            {
                if (credentials != null)
                    await auth.SignOut(credentials);
            }
        }


        private static AWD_User Return_Example_Template_User()
        {
            return new AWD_User()
            {
                Username = "ZHITEST1",
                Alias = "ZHITEST1",
                Country = "1",
                Work_Select = "1",
                Work_Group = "AWD FACIL",
                First_Name = "BATCH",
                Last_Name = "TEST",
                Status = AccountService.Status.Available,
            };
        }

        private static AWD_User Quick_Create_Template_User(string username)
        {
            return new AWD_User()
            {
                Username = username,
                Alias = username,
                Country = "1",
                Work_Group = "AWD FACIL",
                Status = AccountService.Status.Available,
            };
        }

        private static List<AWD_User> Read_And_Get_Users_From_Spreadsheet(string xlsx_file_path)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream file = new FileStream(xlsx_file_path, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(file);
                }

                var sheet = workbook.GetSheetAt(0);
                if (sheet == null)
                {
                    Console.WriteLine($"Could not find spreadsheet at {xlsx_file_path}");
                    return null;
                }

                int name_index = 0;
                var row = sheet.GetRow(0);

                var row_i = 0;
                var awd_users_to_create = new List<AWD_User>();
                var consecutively_skipped_rows = 0;

                while ((row = sheet.GetRow(row_i)) != null)
                {
                    if (!RowHasValues(row)) // skip this row if it doesn't have any values
                    {
                        // if we've skipped more than 20 rows in a row, then we should probably stop reading from this spreadsheet
                        if (consecutively_skipped_rows > 20)
                            break;

                        // keep count of the number of consecutive rows we've skipped
                        consecutively_skipped_rows++;
                        continue;
                    }
                    else
                    {
                        consecutively_skipped_rows = 0;
                    }

                    string username = row.GetCell(name_index).ToString();

                    awd_users_to_create.Add(Quick_Create_Template_User(username));

                    row_i++;
                }
                return awd_users_to_create;
            }
            finally
            {
                workbook?.Close();
            }
        }
        private static bool RowHasValues(IRow row)
        {
            foreach (var i in row.Cells) // iterate over cells
            {
                if (!string.IsNullOrWhiteSpace(i.StringCellValue))
                    return true; // if any cell is not empty, then return true
            }

            return false;
        }

        private static async Task Setup_Env()
        {
            awd_url = "https://awdwebdev.co.ihc.com/awdServer/awd/";
            auth_username = "AWDWBSRV";
            template_username = "AWDWBSRV";

            logger = NLog.LogManager.GetCurrentClassLogger();
            wrapped_logger = new SH_NLogger(logger);

            request = new RequestService(awd_url, wrapped_logger);
            auth = new AuthenticationService(request, wrapped_logger);
            account = new AccountService(request, wrapped_logger);

            var auth_user_password = await auth.GetCyberarkPassword(auth_username, "DEV01");
            credentials = await auth.SignIn(auth_username, auth_user_password);

            client = new HttpClient();

            Console.WriteLine("Signed in");
        }

        static async Task Start_Cloning()
        {
            Console.WriteLine("Setup Complete");
            Credentials? credentials = null;

            var (result, template_userF) = await account.GetUserAccountDetail(template_username, credentials);

            if (!result)
            {
                Console.WriteLine("Failed to get user account details for template user");
                return;
            }

            // TODO write about template user
            Console.WriteLine($"Template user: {template_userF}");
            Console.WriteLine($"\t{template_userF}");

            //AWD_User template_user = null; // TODO
            //await Clone_User(template_user, "ZHITEST1");
        }

        private static async Task<bool> Create_User(AWD_User template_user)
        {
            Console.WriteLine($"Creating user {template_user.Username}");
            var (success, result) = await account.CreateUserAccount(template_user.Username, template_user.First_Name, template_user.Middle_Name, template_user.Last_Name, template_user.Alias,
                template_user.Business_Area, template_user.Work_Select, template_user.Work_Group, template_user.Security_Level, template_user.Redirect,
                template_user.Country, template_user.In_By, template_user.Out_By, template_user.Phone, template_user.Status,
                template_user.Forward_Queue, template_user.Work_Action, template_user.Personal_Queue,
                credentials);

            if (!success)
            {
                Console.WriteLine($"Failed to create user {template_user.Username}, {result}");
            }
            else
            {
                Console.WriteLine($"Successively created user {template_user.Username}");
            }

            return success;
        }
    }
}
