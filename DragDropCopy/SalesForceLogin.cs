using DragDropSF.sforce;
using System;
using System.Web.Services.Protocols;

namespace DragDropSF
{
    public class SalesForceLogin
    {
        private LoginResult lr;
        private SforceService binding;
        private Form1 form1;
        private bool isLoggedIn;
        private string targetPath;

        public SalesForceLogin(Form1 form1, string targetPath) 
        {
            this.form1 = form1;
            this.isLoggedIn = false;
            this.targetPath = targetPath;
        }

        public void run()
        {
            // Make a login call 
            if (login())
            {
                // Set Login Flag to True
                this.isLoggedIn = true;
            }
        }

        private bool login()
        {
            string username = form1.getUsername();
            string password = form1.getPassword();
            this.isLoggedIn = false;

            // Create a service object 
            this.binding = new SforceService();

            // Timeout after a minute 
            this.binding.Timeout = 60000;

            // Try logging in   

            try
            {
                form1.updateTextBox4("\nLogging in...\n");
                this.lr = binding.login(username, password);
            }

            catch (SoapException e)
            {
                form1.updateTextBox4(e.Code.ToString());
                form1.appendTextBox5("An unexpected error has occurred: " + e.Message);
                form1.appendTextBox5(e.StackTrace);

                // Return False to indicate that the login was not successful 
                return false;
            }

            // Check if the password has expired 
            if (lr.passwordExpired)
            {
                form1.updateTextBox4("An error has occurred. Your password has expired.");
                return false;
            }

            // Save old authentication end point URL
            String authEndPoint = binding.Url;
            // Set returned service endpoint URL
            binding.Url = lr.serverUrl;

            /** The sample client application now has an instance of the SforceService
             * that is pointing to the correct endpoint. Next, the sample client
             * application sets a persistent SOAP header (to be included on all
             * subsequent calls that are made with SforceService) that contains the
             * valid sessionId for our login credentials. To do this, the sample
             * client application creates a new SessionHeader object and persist it to
             * the SforceService. Add the session ID returned from the login to the
             * session header
             */
            this.binding.SessionHeaderValue = new SessionHeader();
            this.binding.SessionHeaderValue.sessionId = this.lr.sessionId;

            printUserInfo(lr, authEndPoint);

            // Return true to indicate that we are logged in, pointed  
            // at the right URL and have our security token in place.     
            return true;
        }

        private void printUserInfo(LoginResult lr, String authEP)
        {
            try
            {
                GetUserInfoResult userInfo = lr.userInfo;

                form1.updateTextBox4("\nLogging in ...\n");
                form1.updateTextBox4("Logged in to SalesForce.com as:");
                form1.appendTextBox5("UserID: " + userInfo.userId);
                form1.appendTextBox5("User Full Name: " +
                    userInfo.userFullName);
                form1.appendTextBox5("User Email: " +
                    userInfo.userEmail);
                form1.appendTextBox5("SessionID: " +
                    lr.sessionId);
                form1.appendTextBox5("Auth End Point: " +
                    authEP);
                form1.appendTextBox5("Service End Point: " +
                    lr.serverUrl);
            }
            catch (SoapException e)
            {
                form1.updateTextBox4("An unexpected error has occurred: " + e.Message +
                    " Stack trace: " + e.StackTrace);
            }
        }

        public void logout()
        {
            try
            {
                binding.logout();
                form1.updateTextBox5("");
                form1.updateTextBox4("Logged out.");
            }
            catch (SoapException e)
            {
                // Write the fault message to the console 
                form1.updateTextBox4("An unexpected error has occurred: " + e.Message);
            }
        }
        
        public LoginResult getLoginResult()
        {
            return this.lr;
        }

        public bool getLoginFlag()
        {
            return this.isLoggedIn;
        }
    }
}