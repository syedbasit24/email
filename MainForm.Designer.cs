using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using System.Net.Mail;

namespace EmailMS
{
    partial class MainForm : Form
    {
        private ImapClient imapClient;
        public MainForm()
        {
            InitializeComponent();
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            // Initialize your form and UI elements here.
        }

        private void connectButton_Click(object sender, EventArgs e)
        {
            // Read email server settings from the UI.
            string server = serverTextBox.Text.Trim();
            int port = int.Parse(portTextBox.Text.Trim());
            string username = usernameTextBox.Text.Trim();
            string password = passwordTextBox.Text;

            try
            {
                imapClient = new ImapClient();

                // Connect to the IMAP server.
                imapClient.Connect(server, port, SecureSocketOptions.Auto);

                // Authenticate using the provided username and password.
                imapClient.Authenticate(username, password);

                // Enable features such as IDLE and SORT if needed.
                // imapClient.Capabilities should help you check for supported features.

                // Display a success message or update UI accordingly.
                MessageBox.Show("Connected to the IMAP server successfully.", "Connection Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Handle connection/authentication errors and display an error message.
                MessageBox.Show($"Error: {ex.Message}", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        

    }

    private void readEmailButton_Click(object sender, EventArgs e)
        {
            if (imapClient == null || !imapClient.IsConnected)
            {
                MessageBox.Show("Not connected to the IMAP server.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Open the inbox folder.
                var inbox = imapClient.Inbox;
                inbox.Open(FolderAccess.ReadOnly);

                // Get a list of emails in the inbox.
                var emailList = inbox.Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.Envelope);

                // Clear the ListView.
                emailListView.Items.Clear();

                // Display emails in the ListView.
                foreach (var emailSummary in emailList)
                {
                    var item = new ListViewItem(emailSummary.Envelope.Subject);
                    item.SubItems.Add(emailSummary.Envelope.Date.ToString("yyyy-MM-dd HH:mm:ss"));
                    item.SubItems.Add(emailSummary.Envelope.From.ToString());

                    emailListView.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                // Handle any errors.
                MessageBox.Show($"Error fetching emails: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        
    }

        private void deleteEmailButton_Click(object sender, EventArgs e)
        {
            if (emailListView.SelectedItems.Count == 0)
            {
                MessageBox.Show("Select one or more emails to delete.", "No Emails Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show("Are you sure you want to delete the selected email(s)?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    // Iterate through selected emails and delete them.
                    foreach (ListViewItem item in emailListView.SelectedItems)
                    {
                        // Get the email's subject (you can use other email attributes for identification).
                        string emailSubject = item.Text;

                        // Find the corresponding email in the IMAP folder and mark it for deletion.
                        var uids = imapClient.Inbox.Search(SearchQuery.SubjectContains(emailSubject));
                        if (uids.Count > 0)
                        {
                            imapClient.Inbox.AddFlags(uids, MessageFlags.Deleted, true);
                        }
                    }

                    // Expunge the folder to permanently delete marked emails.
                    imapClient.Inbox.Expunge();

                    // Refresh the email list after deletion.
                    FetchEmailsFromInbox();

                    MessageBox.Show("Selected email(s) deleted successfully.", "Deletion Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    // Handle any errors during the deletion process.
                    MessageBox.Show($"Error deleting email(s): {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            }

        private void searchButton_Click(object sender, EventArgs e)
        {
            // Read the subject entered by the user.
            string searchSubject = searchTextBox.Text.Trim();

            if (string.IsNullOrEmpty(searchSubject))
            {
                MessageBox.Show("Please enter a subject to search for.", "Empty Subject", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Search for emails with the specified subject in the inbox.
                var results = imapClient.Inbox.Search(SearchQuery.SubjectContains(searchSubject));

                // Clear the ListView.
                emailListView.Items.Clear();

                // Display search results in the ListView.
                foreach (var result in results)
                {
                    var message = imapClient.Inbox.GetMessage(result);

                    var item = new ListViewItem(message.Subject);
                    item.SubItems.Add(message.Date.ToString("yyyy-MM-dd HH:mm:ss"));
                    item.SubItems.Add(message.From.ToString());

                    emailListView.Items.Add(item);
                }

                if (results.Count == 0)
                {
                    MessageBox.Show("No emails found with the specified subject.", "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                // Handle any errors during the search process.
                MessageBox.Show($"Error searching emails: {ex.Message}", "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void moveEmailButton_Click(object sender, EventArgs e)
        {
            // Read the selected source and destination folders.
            string sourceFolder = sourceFolderComboBox.SelectedItem?.ToString();
            string destinationFolder = destinationFolderComboBox.SelectedItem?.ToString();

            if (string.IsNullOrEmpty(sourceFolder) || string.IsNullOrEmpty(destinationFolder))
            {
                MessageBox.Show("Please select both source and destination folders.", "Missing Folders", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Check if any emails are selected in the ListView.
            if (emailListView.SelectedItems.Count == 0)
            {
                MessageBox.Show("Select one or more emails to move.", "No Emails Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show($"Are you sure you want to move the selected email(s) from '{sourceFolder}' to '{destinationFolder}'?", "Confirm Move", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    // Iterate through selected emails and move them to the destination folder.
                    foreach (ListViewItem item in emailListView.SelectedItems)
                    {
                        // Get the email's subject (you can use other email attributes for identification).
                        string emailSubject = item.Text;

                        // Find the corresponding email in the source folder and move it.
                        var uids = imapClient.Inbox.Search(SearchQuery.SubjectContains(emailSubject));
                        if (uids.Count > 0)
                        {
                            imapClient.Inbox.MoveTo(uids, imapClient.GetFolder(destinationFolder));
                        }
                    }

                    // Refresh the email list after moving.
                    FetchEmailsFromInbox();

                    MessageBox.Show("Selected email(s) moved successfully.", "Move Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    // Handle any errors during the move process.
                    MessageBox.Show($"Error moving email(s): {ex.Message}", "Move Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        private void flagEmailButton_Click(object sender, EventArgs e)
        {
            // Check if any emails are selected in the ListView.
            if (emailListView.SelectedItems.Count == 0)
            {
                MessageBox.Show("Select one or more emails to flag.", "No Emails Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Iterate through selected emails and flag them.
                foreach (ListViewItem item in emailListView.SelectedItems)
                {
                    // Get the email's subject (you can use other email attributes for identification).
                    string emailSubject = item.Text;

                    // Find the corresponding email in the IMAP folder and flag it.
                    var uids = imapClient.Inbox.Search(SearchQuery.SubjectContains(emailSubject));
                    if (uids.Count > 0)
                    {
                        imapClient.Inbox.AddFlags(uids, MessageFlags.Flagged, true);
                    }
                }

                // Refresh the email list after flagging.
                FetchEmailsFromInbox();

                MessageBox.Show("Selected email(s) flagged successfully.", "Flagging Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Handle any errors during the flagging process.
                MessageBox.Show($"Error flagging email(s): {ex.Message}", "Flagging Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Text = "MainForm";
        }

        #endregion
    }
}