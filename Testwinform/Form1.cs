using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Testwinform
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
       
        }

//new starter form
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

             textBox3.Text = textBox2.Text + ", " + textBox1.Text;
             textBox17.Text = textBox1.Text + " " + textBox2.Text;
             textBox6.Text = textBox1.Text + " " + textBox2.Text;
             
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
             textBox3.Text = textBox2.Text + ", " + textBox1.Text;
             textBox17.Text = textBox1.Text + " " + textBox2.Text;
             textBox6.Text = textBox1.Text + " " + textBox2.Text;
             
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            textBox4.Text = dateTimePicker1.Value.ToShortDateString();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            textBox11.Text = "\\\\Sws.int.southernwater.co.uk\\sw-dfs-00\\Worthing-users\\" + textBox7.Text;
            textBox15.Text = textBox7.Text;
            textBox37.Text = textBox7.Text;
            textBox38.Text = textBox7.Text;

        }
        // new starter H drive
        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(comboBox2.Text))
            {
                MessageBox.Show("Please select Home Drive", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox11.Text)) 
            {
                MessageBox.Show("Please fill newstarter form", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
                textBox12.Text = "Set-ADUser " + textBox7.Text + " -homedirectory \"" + textBox11.Text + "\" -homedrive " + comboBox2.Text + ":";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox12.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox12.Text);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox12.Text = "";
            textBox11.Text = "";
        }

        // new starter account
        private void button4_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text) && string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Please enter the First Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Please enter the Surname", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }   
            else if (string.IsNullOrEmpty(comboBox1.Text))
            {
                MessageBox.Show("Please select Logon Script", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox6.Text))
            {
                MessageBox.Show("Please enter the Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox10.Text))
            {
                MessageBox.Show("Please enter the Company Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
           
           else if (string.IsNullOrEmpty(textBox7.Text))
            {
                MessageBox.Show("Please Enter Logon Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (textBox7.Text.Length > 8)
            {
                MessageBox.Show("Logon Name must be 7 or 8 characters long", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (textBox7.Text.Length < 7)
            {
                MessageBox.Show("Logon Name must be 7 or 8 characters long", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(comboBox3.Text))
            {
                MessageBox.Show("Please select Path", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }

            else
            {
                if (textBox4.Text.Equals(""))
                {
                    DateTime date = this.dateTimePicker1.Value;
                    this.textBox4.Text = date.ToShortDateString();
                }
                if (string.IsNullOrEmpty(textBox5.Text) && string.IsNullOrEmpty(textBox8.Text) && string.IsNullOrEmpty(textBox9.Text))
                {
                    textBox13.Text = "New-ADUser -SamAccountName " + textBox7.Text + " -GivenName " + textBox1.Text + " -Surname " + textBox2.Text + " -DisplayName \"" + textBox3.Text + "\" -Path " + comboBox3.Text + " -AccountExpirationDate " + textBox4.Text + " -Company \"" + textBox10.Text + "\" -Description \"" + textBox14.Text + "\" -Name \"" + textBox6.Text + "\" -ScriptPath " + comboBox1.Text + " -ChangePasswordAtLogon $true" + " -UserPrincipalName \"" + textBox7.Text + "@southernwater.co.uk\" -AccountPassword (Read-Host -AsSecureString \"AccountPassword\") -Enabled $true";
                }
                else if (string.IsNullOrEmpty(textBox5.Text) && string.IsNullOrEmpty(textBox9.Text))
                {
                    textBox13.Text = "New-ADUser -SamAccountName " + textBox7.Text + " -GivenName " + textBox1.Text + " -Surname " + textBox2.Text + " -DisplayName \"" + textBox3.Text + "\" -Path " + comboBox3.Text + " -AccountExpirationDate " + textBox4.Text + " -Title " + textBox8.Text + " -Company \"" + textBox10.Text + "\" -Description \"" + textBox14.Text + "\" -Name \"" + textBox6.Text + "\" -ScriptPath " + comboBox1.Text + " -ChangePasswordAtLogon $true" + " -UserPrincipalName \"" + textBox7.Text + "@southernwater.co.uk\" -AccountPassword (Read-Host -AsSecureString \"AccountPassword\") -Enabled $true";
                }
                else if (string.IsNullOrEmpty(textBox5.Text) && string.IsNullOrEmpty(textBox8.Text))
                {
                    textBox13.Text = "New-ADUser -SamAccountName " + textBox7.Text + " -GivenName " + textBox1.Text + " -Surname " + textBox2.Text + " -DisplayName \"" + textBox3.Text + "\" -Path " + comboBox3.Text + " -AccountExpirationDate " + textBox4.Text + " -Department " + textBox9.Text + " -Company \"" + textBox10.Text + "\" -Description \"" + textBox14.Text + "\" -Name \"" + textBox6.Text + "\" -ScriptPath " + comboBox1.Text + " -ChangePasswordAtLogon $true" + " -UserPrincipalName \"" + textBox7.Text + "@southernwater.co.uk\" -AccountPassword (Read-Host -AsSecureString \"AccountPassword\") -Enabled $true";
                }
                else if (string.IsNullOrEmpty(textBox8.Text) && string.IsNullOrEmpty(textBox9.Text))
                {
                    textBox13.Text = "New-ADUser -SamAccountName " + textBox7.Text + " -GivenName " + textBox1.Text + " -Surname " + textBox2.Text + " -DisplayName \"" + textBox3.Text + "\" -Path " + comboBox3.Text + " -AccountExpirationDate " + textBox4.Text + " -Office " + textBox5.Text + " -Company \"" + textBox10.Text + "\" -Description \"" + textBox14.Text + "\" -Name \"" + textBox6.Text + "\" -ScriptPath " + comboBox1.Text + " -ChangePasswordAtLogon $true" + " -UserPrincipalName \"" + textBox7.Text + "@southernwater.co.uk\" -AccountPassword (Read-Host -AsSecureString \"AccountPassword\") -Enabled $true";
                }
                else if (string.IsNullOrEmpty(textBox5.Text))
                {
                    textBox13.Text = "New-ADUser -SamAccountName " + textBox7.Text + " -GivenName " + textBox1.Text + " -Surname " + textBox2.Text + " -DisplayName \"" + textBox3.Text + "\" -Path " + comboBox3.Text + " -AccountExpirationDate " + textBox4.Text + " -Title " + textBox8.Text + " -Department " + textBox9.Text + " -Company \"" + textBox10.Text + "\" -Description \"" + textBox14.Text + "\" -Name \"" + textBox6.Text + "\" -ScriptPath " + comboBox1.Text + " -ChangePasswordAtLogon $true" + " -UserPrincipalName \"" + textBox7.Text + "@southernwater.co.uk\" -AccountPassword (Read-Host -AsSecureString \"AccountPassword\") -Enabled $true";
                }
                else if (string.IsNullOrEmpty(textBox8.Text))
                {
                    textBox13.Text = "New-ADUser -SamAccountName " + textBox7.Text + " -GivenName " + textBox1.Text + " -Surname " + textBox2.Text + " -DisplayName \"" + textBox3.Text + "\" -Path " + comboBox3.Text + " -AccountExpirationDate " + textBox4.Text + " -Office " + textBox5.Text + " -Department " + textBox9.Text + " -Company \"" + textBox10.Text + "\" -Description \"" + textBox14.Text + "\" -Name \"" + textBox6.Text + "\" -ScriptPath " + comboBox1.Text + " -ChangePasswordAtLogon $true" + " -UserPrincipalName \"" + textBox7.Text + "@southernwater.co.uk\" -AccountPassword (Read-Host -AsSecureString \"AccountPassword\") -Enabled $true";
                }
                else if (string.IsNullOrEmpty(textBox9.Text))
                {
                    textBox13.Text = "New-ADUser -SamAccountName " + textBox7.Text + " -GivenName " + textBox1.Text + " -Surname " + textBox2.Text + " -DisplayName \"" + textBox3.Text + "\" -Path " + comboBox3.Text + " -AccountExpirationDate " + textBox4.Text + " -Office " + textBox5.Text + " -Title " + textBox8.Text + " -Company \"" + textBox10.Text + "\" -Description \"" + textBox14.Text + "\" -Name \"" + textBox6.Text + "\" -ScriptPath " + comboBox1.Text + " -ChangePasswordAtLogon $true" + " -UserPrincipalName \"" + textBox7.Text + "@southernwater.co.uk\" -AccountPassword (Read-Host -AsSecureString \"AccountPassword\") -Enabled $true";
                }
                else
                {
                    textBox13.Text = "New-ADUser -SamAccountName " + textBox7.Text + " -GivenName " + textBox1.Text + " -Surname " + textBox2.Text + " -DisplayName \"" + textBox3.Text + "\" -Path " + comboBox3.Text + " -AccountExpirationDate " + textBox4.Text + " -Office " + textBox5.Text + " -Title " + textBox8.Text + " -Department " + textBox9.Text + " -Company \"" + textBox10.Text + "\" -Description \"" + textBox14.Text + "\" -Name \"" + textBox6.Text + "\" -ScriptPath " + comboBox1.Text + " -ChangePasswordAtLogon $true" + " -UserPrincipalName \"" + textBox7.Text + "@southernwater.co.uk\" -AccountPassword (Read-Host -AsSecureString \"AccountPassword\") -Enabled $true";
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox13.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox13.Text);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            textBox13.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
        }
        // new starter mailbox
        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox16.Text = comboBox10.Text + textBox17.Text + "\"";
        }
        private void button7_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox15.Text))
            {
                MessageBox.Show("Please Enter Logon Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
           
            else if (string.IsNullOrEmpty(textBox17.Text))
            {
                MessageBox.Show("Please Enter Full Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(comboBox10.Text))
            {
                MessageBox.Show("Please Select Path", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(comboBox4.Text))
            {
                MessageBox.Show("Please Select Storage Group", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            
            else
            {
                textBox18.Text = "Enable-Mailbox -Identity " + textBox16.Text + " -Alias \"" + textBox15.Text + "\" -database " + comboBox4.Text + "; Set-CASMailbox -Identity \"" + textBox15.Text + "\" -ActiveSyncEnabled $false; " + "Set-Mailbox -Identity \"" + textBox15.Text + "\" -UseDatabaseQuotaDefaults $false";
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox18.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox18.Text);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            textBox18.Text = "";
            textBox17.Text = "";
            textBox16.Text = "";
             textBox15.Text = "";
             comboBox10.Text= "";
             comboBox4.Text = "";
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            textBox21.Text = textBox20.Text + comboBox7.Text;
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox21.Text = textBox20.Text + comboBox7.Text;
        }
// shared mailbox
        private void button10_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox19.Text))
            {
                MessageBox.Show("Please Enter MailBox Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (textBox19.Text.Length > 20)
            {
                MessageBox.Show("MailBox Name must be less than 20 characters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }            
            else if (string.IsNullOrEmpty(textBox20.Text))
            {
                MessageBox.Show("Please Enter Mail Address", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (textBox20.Text.Length > 20)
            {
                MessageBox.Show("Logon Name must be less than 20 characters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            
            else if (string.IsNullOrEmpty(comboBox7.Text))
            {
                MessageBox.Show("Please Select Domain Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(comboBox5.Text))
            {
                MessageBox.Show("Please Select Storage Group", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            
            else if (string.IsNullOrEmpty(textBox21.Text))
            {
                MessageBox.Show("The Please enter Mailbox Address", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            //else if (string.IsNullOrEmpty(textBox34.Text))
            //{
              //  MessageBox.Show("Please the enter Password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            //}
            //else if (textBox34.Text.Length < 8)
            //{
              //  MessageBox.Show("Password must be at least 8 characters long", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            //}
            else 
            {
                textBox22.Text = "New-Mailbox -Name " + textBox19.Text + " -Alias " + textBox20.Text + " -OrganizationalUnit 'sws.int.southernwater.co.uk/Users_New_Structure/OtherUsers/Exchange_Shared_Mailboxs'" + " -UserPrincipalName " + textBox21.Text + " -SamAccountName " + textBox20.Text + " -Password (Read-Host -AsSecureString \"Password\") -ResetPasswordOnNextLogon $false" + " -database " + comboBox5.Text + "; Set-CASMailbox -Identity \"" + textBox20.Text + "\" -ActiveSyncEnabled $false; Set-Mailbox -Identity \"" + textBox20.Text + "\" -UseDatabaseQuotaDefaults $false;";
            }
           
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox22.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox22.Text);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            textBox22.Text = "";
            textBox19.Text = "";
            textBox20.Text = "";
            textBox21.Text = "";
            comboBox5.Text = "";
            
        }
//shared mailbox send as permission

        private void textBox23_TextChanged(object sender, EventArgs e)
        {
            textBox26.Text = textBox23.Text;
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            textBox25.Text = textBox24.Text;
        }

        private void textBox34_TextChanged(object sender, EventArgs e)
        {
            textBox46.Text = textBox34.Text;
        }

        private void textBox43_TextChanged(object sender, EventArgs e)
        {
            textBox47.Text = textBox43.Text;
        }

        private void textBox44_TextChanged(object sender, EventArgs e)
        {
            textBox48.Text = textBox44.Text;
        }

        private void textBox45_TextChanged(object sender, EventArgs e)
        {
            textBox49.Text = textBox45.Text;
        }
               private void button13_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox23.Text))
            {
                MessageBox.Show("Please Enter MailBox Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox24.Text))
            {
                MessageBox.Show("Please Enter atleast one Full Name/Logon Name to give Send As Permission", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if ((string.IsNullOrEmpty(textBox24.Text)) && (string.IsNullOrEmpty(textBox34.Text)) && (string.IsNullOrEmpty(textBox43.Text)) && (string.IsNullOrEmpty(textBox44.Text)) && (string.IsNullOrEmpty(textBox45.Text)))
            {
                MessageBox.Show("Please Enter atleast one Full Name/Logon Name to give Send As Permission", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if ((string.IsNullOrEmpty(textBox34.Text)) && (string.IsNullOrEmpty(textBox43.Text)) && (string.IsNullOrEmpty(textBox44.Text)) && (string.IsNullOrEmpty(textBox45.Text)))
            {
                textBox27.Text = "Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox24.Text + "\" -ExtendedRights \"Send-As\""; 
            }
            else if ((string.IsNullOrEmpty(textBox43.Text)) && (string.IsNullOrEmpty(textBox44.Text)) && (string.IsNullOrEmpty(textBox45.Text)))
            {
                textBox27.Text = "Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox24.Text + "\" -ExtendedRights \"Send-As\";" + " Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox34.Text + "\" -ExtendedRights \"Send-As\"";
            }
            else if ((string.IsNullOrEmpty(textBox44.Text)) && (string.IsNullOrEmpty(textBox45.Text)))
            {
                textBox27.Text = "Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox24.Text + "\" -ExtendedRights \"Send-As\";" + " Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox34.Text + "\" -ExtendedRights \"Send-As\";" + " Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox43.Text + "\" -ExtendedRights \"Send-As\"";
            }
            else if ((string.IsNullOrEmpty(textBox45.Text)))
            {
                textBox27.Text = "Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox24.Text + "\" -ExtendedRights \"Send-As\";" + " Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox34.Text + "\" -ExtendedRights \"Send-As\";" + " Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox43.Text + "\" -ExtendedRights \"Send-As\";" + " Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox44.Text + "\" -ExtendedRights \"Send-As\"";
            }
            else
            {
                textBox27.Text = "Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox24.Text + "\" -ExtendedRights \"Send-As\";" + " Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox34.Text + "\" -ExtendedRights \"Send-As\";" + " Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox43.Text + "\" -ExtendedRights \"Send-As\";" + " Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox44.Text + "\" -ExtendedRights \"Send-As\";" + " Add-ADPermission -Identity \"" + textBox23.Text + "\" -User \"" + textBox45.Text + "\" -ExtendedRights \"Send-As\"";
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox27.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox27.Text);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            textBox27.Text = "";
            textBox23.Text = "";
            textBox24.Text = "";
            textBox34.Text = "";
            textBox43.Text = "";
            textBox44.Text = "";
            textBox45.Text = "";
        }
// shared mailbox full access
        private void button16_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox26.Text))
            {
                MessageBox.Show("Please Enter MailBox Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox25.Text))
            {
                MessageBox.Show("Please Enter atleast one Full Name/Logon Name to give Full Access Permission", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if ((string.IsNullOrEmpty(textBox25.Text)) && (string.IsNullOrEmpty(textBox46.Text)) && (string.IsNullOrEmpty(textBox47.Text)) && (string.IsNullOrEmpty(textBox48.Text)) && (string.IsNullOrEmpty(textBox49.Text)))
            {
                MessageBox.Show("Please Enter atleast one Full Name/Logon Name to give Full Access Permission", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if ((string.IsNullOrEmpty(textBox46.Text)) && (string.IsNullOrEmpty(textBox47.Text)) && (string.IsNullOrEmpty(textBox48.Text)) && (string.IsNullOrEmpty(textBox49.Text)))
            {
                textBox28.Text = "Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox25.Text + "\" -AccessRights FullAccess"; 
            }
            else if ((string.IsNullOrEmpty(textBox47.Text)) && (string.IsNullOrEmpty(textBox48.Text)) && (string.IsNullOrEmpty(textBox49.Text)))
            {
                textBox28.Text = "Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox25.Text + "\" -AccessRights FullAccess;" + " Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox46.Text + "\" -AccessRights FullAccess";
            }
            else if ((string.IsNullOrEmpty(textBox48.Text)) && (string.IsNullOrEmpty(textBox49.Text)))
            {
                textBox28.Text = "Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox25.Text + "\" -AccessRights FullAccess;" + " Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox46.Text + "\" -AccessRights FullAccess;" + " Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox47.Text + "\" -AccessRights FullAccess";
            }
            else if ((string.IsNullOrEmpty(textBox49.Text)))
            {
                textBox28.Text = "Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox25.Text + "\" -AccessRights FullAccess;" + " Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox46.Text + "\" -AccessRights FullAccess;" + " Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox47.Text + "\" -AccessRights FullAccess;" + " Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox48.Text + "\" -AccessRights FullAccess";
            }
            else
            {
                textBox28.Text = "Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox25.Text + "\" -AccessRights FullAccess;" + " Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox46.Text + "\" -AccessRights FullAccess;" + " Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox47.Text + "\" -AccessRights FullAccess;" + " Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox48.Text + "\" -AccessRights FullAccess;" + " Add-MailboxPermission -Identity \"" + textBox26.Text + "\" -User \"" + textBox49.Text + "\" -AccessRights FullAccess";
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox28.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox28.Text);
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            textBox28.Text = "";
            textBox26.Text = "";
            textBox25.Text = "";
            textBox46.Text = "";
            textBox47.Text = "";
            textBox48.Text = "";
            textBox49.Text = "";
        }
        // distribution group
        private void button19_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox29.Text))
            {
                MessageBox.Show("Please Enter Group Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox32.Text))
            {
                MessageBox.Show("Please Enter Owner Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox30.Text))
            {
                MessageBox.Show("Please Enter Logon/Alias Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(comboBox9.Text))
            {
                MessageBox.Show("Please Select Type", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(comboBox8.Text))
            {
                MessageBox.Show("Please Select Path", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            
            else
            {
                textBox31.Text = "New-DistributionGroup -Name " + textBox29.Text + " -OrganizationalUnit " + comboBox8.Text + " -SAMAccountName " + textBox30.Text + " -Type " + comboBox9.Text + " -ManagedBy \"" + textBox32.Text + "\"";
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox31.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox31.Text);
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            textBox31.Text = "";
            textBox29.Text = "";
            textBox32.Text = "";
            textBox30.Text = "";
            
        }
        //new starter member of
        private void button22_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox37.Text))
            {
                MessageBox.Show("Please Enter Logon Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox35.Text = "\"Access_All SW users\", \"Access_iWelcome_AllUsers\", \"AllowAccess-SWWallpaper-New\", \"S1_Archive_Group\", \"SW Outlook 2003 PST Restriction\", \"SW.APPSENSE.APPDATA\", \"SW.APPSENSE.ENV\", \"SW.APPSENSE.FLEXSTOP\", \"SW.CXA.MYOFFICE2.PROD\", \"SW.CXA.USERS.PROD\", \"UNI_SWPRN-UNI\" | Add-ADGroupMember -Member \"" + textBox37.Text + "\"";
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox35.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox35.Text);
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            textBox35.Text = "";
            textBox37.Text = "";
            
        }
        //new starter member of 
        private void button25_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox38.Text))
            {
                MessageBox.Show("Please Enter Logon Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox36.Text = "\"UNI_SWPRN-UNI\", \"S1_Archive_Group\", \"SW.APPSENSE.APPDATA\", \"SW.APPSENSE.ENV\", \"SW.APPSENSE.FLEXSTOP\", \"SW.CPS.CC.MainGuideHelp.PROD\", \"SW.CPS.CC.WinExplorerGen.Prod\", \"SW.CPS.CCUsers\",  \"SW.CPS.CC.AcrobatReader.Prod\", \"SW.CPS.CC.MicrosoftOffice2003.Prod\", \"SW.CPS.CC.MicrosoftOutlook.Prod\", \"SW.CPS.CC.Amyuni.Prod\", \"SW_Internet_Deny_SCS\", \"WQM_Agent_U\", \"WQMScript\", \"SW Outlook 2003 PST Restriction\", \"sw.cxa.gdrive.prod\",  \"sw.cxa.myoffice2.prod\", \"sw.cxa.users.prod\" | Add-ADGroupMember -Member \"" + textBox38.Text + "\"";
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox36.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox36.Text);
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            textBox38.Text = "";
            textBox36.Text = "";

        }
        // add distribution group memebers
        private void button28_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox42.Text))
            {
                MessageBox.Show("Please Enter Distribution list to add members", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox33.Text))
            {
                MessageBox.Show("Please Enter atleast one member to add", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox33.Text) && string.IsNullOrEmpty(textBox39.Text) && string.IsNullOrEmpty(textBox40.Text) && string.IsNullOrEmpty(textBox50.Text) && string.IsNullOrEmpty(textBox51.Text))
            {
                MessageBox.Show("Please Enter atleast one member to add", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox39.Text) && string.IsNullOrEmpty(textBox40.Text) && string.IsNullOrEmpty(textBox50.Text) && string.IsNullOrEmpty(textBox51.Text))
            {
                textBox41.Text = "Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox33.Text + "\"";
            }
            else if (string.IsNullOrEmpty(textBox40.Text) && string.IsNullOrEmpty(textBox50.Text) && string.IsNullOrEmpty(textBox51.Text))
            {
                textBox41.Text = "Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox33.Text + "\"" + "; Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox39.Text + "\";";
            }
            else if (string.IsNullOrEmpty(textBox50.Text) && string.IsNullOrEmpty(textBox51.Text))
            {
                textBox41.Text = "Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox33.Text + "\"" + "; Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox39.Text + "\"" + "; Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox40.Text + "\";";
            }
            else if (string.IsNullOrEmpty(textBox51.Text))
            {
                textBox41.Text = "Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox33.Text + "\"" + "; Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox39.Text + "\"" + "; Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox40.Text + "\"" + "; Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox50.Text + "\";";
            }
            else 
            {
                textBox41.Text = "Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox33.Text + "\"" + "; Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox39.Text + "\"" + "; Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox40.Text + "\"" + "; Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox50.Text + "\"" + "; Add-DistributionGroupMember -Identity \"" + textBox42.Text + "\" -Member \"" + textBox51.Text + "\";";
            }

        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {
            textBox42.Text = textBox30.Text;
        }

        private void button29_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox41.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox41.Text);
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            textBox33.Text = "";
            textBox39.Text = "";
            textBox40.Text = "";
            textBox41.Text = "";
            textBox42.Text = "";
            textBox50.Text = "";
            textBox51.Text = "";
        }

        //remove shared mailbox access
        private void textBox58_TextChanged(object sender, EventArgs e)
        {
            textBox65.Text = textBox58.Text;
        }

        private void textBox57_TextChanged(object sender, EventArgs e)
        {
            textBox64.Text = textBox57.Text;
        }

        private void textBox55_TextChanged(object sender, EventArgs e)
        {
            textBox62.Text = textBox55.Text;
        }

        private void textBox54_TextChanged(object sender, EventArgs e)
        {
            textBox61.Text = textBox54.Text;
        }

        private void textBox53_TextChanged(object sender, EventArgs e)
        {
            textBox60.Text = textBox53.Text;
        }

        private void textBox52_TextChanged(object sender, EventArgs e)
        {
            textBox59.Text = textBox52.Text;
        }
        private void button33_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox58.Text))
            {
                MessageBox.Show("Please Enter MailBox Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox57.Text))
            {
                MessageBox.Show("Please Enter atleast one Full Name/Logon Name to Remove Send As Permission", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if ((string.IsNullOrEmpty(textBox57.Text)) && (string.IsNullOrEmpty(textBox55.Text)) && (string.IsNullOrEmpty(textBox54.Text)) && (string.IsNullOrEmpty(textBox53.Text)) && (string.IsNullOrEmpty(textBox52.Text)))
            {
                MessageBox.Show("Please Enter atleast one Full Name/Logon Name to Remove Send As Permission", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if ((string.IsNullOrEmpty(textBox55.Text)) && (string.IsNullOrEmpty(textBox54.Text)) && (string.IsNullOrEmpty(textBox53.Text)) && (string.IsNullOrEmpty(textBox52.Text)))
            {
                textBox56.Text = "Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox57.Text + "\" -ExtendedRights \"Send-As\"";
            }
            else if ((string.IsNullOrEmpty(textBox54.Text)) && (string.IsNullOrEmpty(textBox53.Text)) && (string.IsNullOrEmpty(textBox52.Text)))
            {
                textBox56.Text = "Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox57.Text + "\" -ExtendedRights \"Send-As\";" + " Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox55.Text + "\" -ExtendedRights \"Send-As\"";
            }
            else if ((string.IsNullOrEmpty(textBox53.Text)) && (string.IsNullOrEmpty(textBox52.Text)))
            {
                textBox56.Text = "Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox57.Text + "\" -ExtendedRights \"Send-As\";" + " Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox55.Text + "\" -ExtendedRights \"Send-As\";" + " Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox54.Text + "\" -ExtendedRights \"Send-As\"";
            }
            else if ((string.IsNullOrEmpty(textBox52.Text)))
            {
                textBox56.Text = "Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox57.Text + "\" -ExtendedRights \"Send-As\";" + " Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox55.Text + "\" -ExtendedRights \"Send-As\";" + " Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox54.Text + "\" -ExtendedRights \"Send-As\";" + " Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox53.Text + "\" -ExtendedRights \"Send-As\"";
            }
            else 
            {
                textBox56.Text = "Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox57.Text + "\" -ExtendedRights \"Send-As\";" + " Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox55.Text + "\" -ExtendedRights \"Send-As\";" + " Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox54.Text + "\" -ExtendedRights \"Send-As\";" + " Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox53.Text + "\" -ExtendedRights \"Send-As\";" + " Remove-ADPermission -Identity \"" + textBox58.Text + "\" -User \"" + textBox52.Text + "\" -ExtendedRights \"Send-As\"";
            }

        }

        private void button32_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox56.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox56.Text);
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            textBox58.Text = "";
            textBox57.Text = "";
            textBox56.Text = "";
            textBox55.Text = "";
            textBox54.Text = "";
            textBox53.Text = "";
            textBox52.Text = "";
        }
        //remove full access permission

        private void button36_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox65.Text))
            {
                MessageBox.Show("Please Enter MailBox Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox64.Text))
            {
                MessageBox.Show("Please Enter atleast one Full Name/Logon Name to Remove Full Access Permission", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if ((string.IsNullOrEmpty(textBox64.Text)) && (string.IsNullOrEmpty(textBox62.Text)) && (string.IsNullOrEmpty(textBox61.Text)) && (string.IsNullOrEmpty(textBox60.Text)) && (string.IsNullOrEmpty(textBox59.Text)))
            {
                MessageBox.Show("Please Enter atleast one Full Name/Logon Name to Remove Full Access Permission", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if ((string.IsNullOrEmpty(textBox62.Text)) && (string.IsNullOrEmpty(textBox61.Text)) && (string.IsNullOrEmpty(textBox60.Text)) && (string.IsNullOrEmpty(textBox59.Text)))
            {
                textBox63.Text = "Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox64.Text + "\" -AccessRights FullAccess";
            }
            else if ((string.IsNullOrEmpty(textBox61.Text)) && (string.IsNullOrEmpty(textBox60.Text)) && (string.IsNullOrEmpty(textBox59.Text)))
            {
                textBox63.Text = "Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox64.Text + "\" -AccessRights FullAccess;" + " Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox62.Text + "\" -AccessRights FullAccess";
            }
            else if ((string.IsNullOrEmpty(textBox60.Text)) && (string.IsNullOrEmpty(textBox59.Text)))
            {
                textBox63.Text = "Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox64.Text + "\" -AccessRights FullAccess;" + " Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox62.Text + "\" -AccessRights FullAccess;" + " Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox61.Text + "\" -AccessRights FullAccess";
            }
            else if ((string.IsNullOrEmpty(textBox59.Text)))
            {
                textBox63.Text = "Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox64.Text + "\" -AccessRights FullAccess;" + " Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox62.Text + "\" -AccessRights FullAccess;" + " Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox61.Text + "\" -AccessRights FullAccess;" + " Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox60.Text + "\" -AccessRights FullAccess";
            }
            else
            {
                textBox63.Text = "Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox64.Text + "\" -AccessRights FullAccess;" + " Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox62.Text + "\" -AccessRights FullAccess;" + " Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox61.Text + "\" -AccessRights FullAccess;" + " Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox60.Text + "\" -AccessRights FullAccess;" + " Remove-MailboxPermission -Identity \"" + textBox65.Text + "\" -User \"" + textBox59.Text + "\" -AccessRights FullAccess";
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox63.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox63.Text);
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            textBox65.Text = "";
            textBox64.Text = "";
            textBox62.Text = "";
            textBox61.Text = "";
            textBox60.Text = "";
            textBox59.Text = "";
            textBox63.Text = "";
        }
        //Remove Distribution Group Member
        private void button39_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox68.Text))
            {
                MessageBox.Show("Please Enter Distribution list to Remove members", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox72.Text))
            {
                MessageBox.Show("Please Enter atleast one member to Remove", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox72.Text) && string.IsNullOrEmpty(textBox71.Text) && string.IsNullOrEmpty(textBox70.Text) && string.IsNullOrEmpty(textBox67.Text) && string.IsNullOrEmpty(textBox66.Text))
            {
                MessageBox.Show("Please Enter atleast one member to Remove", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox71.Text) && string.IsNullOrEmpty(textBox70.Text) && string.IsNullOrEmpty(textBox67.Text) && string.IsNullOrEmpty(textBox66.Text))
            {
                textBox69.Text = "Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox72.Text + "\"";
            }
            else if (string.IsNullOrEmpty(textBox70.Text) && string.IsNullOrEmpty(textBox67.Text) && string.IsNullOrEmpty(textBox66.Text))
            {
                textBox69.Text = "Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox72.Text + "\"" + "; Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox71.Text + "\";";
            }
            else if (string.IsNullOrEmpty(textBox67.Text) && string.IsNullOrEmpty(textBox66.Text))
            {
                textBox69.Text = "Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox72.Text + "\"" + "; Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox71.Text + "\"" + "; Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox70.Text + "\";";
            }
            else if (string.IsNullOrEmpty(textBox66.Text))
            {
                textBox69.Text = "Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox72.Text + "\"" + "; Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox71.Text + "\"" + "; Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox70.Text + "\"" + "; Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox67.Text + "\";";
            }
            else
            {
                textBox69.Text = "Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox72.Text + "\"" + "; Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox71.Text + "\"" + "; Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox70.Text + "\"" + "; Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox67.Text + "\"" + "; Remove-DistributionGroupMember -Identity \"" + textBox68.Text + "\" -Member \"" + textBox66.Text + "\";";
            }

        }

        private void button38_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox69.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox69.Text);
            }
        }

        private void button37_Click(object sender, EventArgs e)
        {
            textBox68.Text = "";
            textBox72.Text = "";
            textBox71.Text = "";
            textBox70.Text = "";
            textBox67.Text = "";
            textBox66.Text = "";
            textBox69.Text = "";
        }
 
        //Diable AD user        
        private void button42_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox73.Text))
            {
                MessageBox.Show("Please Enter Logon Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox74.Text = "Get-ADUser " + textBox73.Text + " | Move-ADObject -TargetPath \"OU=Disabled_Users, OU=Disabled_Objects, DC=sws, DC=int, DC=southernwater, DC=co, DC=uk\"; Disable-ADAccount -Identity \"" + textBox73.Text+ "\"";
            }
        }

        private void button41_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox74.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox74.Text);
            }
        }

        private void button40_Click(object sender, EventArgs e)
        {
            textBox73.Text = "";
            textBox74.Text = "";
        }

        private void button45_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox73.Text))
            {
                MessageBox.Show("Please Enter Logon Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox75.Text = "Set-Mailbox -Identity \"" + textBox73.Text + "\" -HiddenFromAddressListsEnabled $true";
            }

        }
        private void button44_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox75.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox75.Text);
            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            textBox73.Text = "";
            textBox75.Text = "";
        }

     //password expiry reports

        private void button49_Click(object sender, EventArgs e)
        {
           textBox76.Text = "Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} –Properties \"SamAccountName\",\"mail\",\"pwdLastSet\",\"msDS-UserPasswordExpiryTimeComputed\" | Select-Object -Property \"SamAccountName\",\"mail\",@{Name=\"Password Last Set\"; Expression={[datetime]::FromFileTime($_.\"pwdLastSet\")}}, @{Name=\"Password Expiry Date\"; Expression={[datetime]::FromFileTime($_.\"msDS-UserPasswordExpiryTimeComputed\")}} | Export-CSV \"H:\\AccountExpiryReport.csv\" -NoTypeInformation -Encoding UTF8";
        }

        private void button48_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox77.Text))
            {
                MessageBox.Show("Please Enter The Days In Number To Genarate The Report", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox76.Text = "Search-ADAccount -AccountExpiring -TimeSpan \"" + textBox77.Text + "\" | Select-Object Name,AccountExpirationDate | Sort-Object AccountExpirationDate | Export-CSV \"H:\\AccountExpiryReport.csv\" -NoTypeInformation -Encoding UTF8";
            }
        }

        private void button47_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(textBox76.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox76.Text);
            }
        }

        private void button46_Click(object sender, EventArgs e)
        {
            textBox76.Text ="";
            textBox77.Text ="";
        }
        //locked out
        private void button75_Click(object sender, EventArgs e)
        {
            textBox76.Text = "Get-ADUser -Filter *  -Properties SamAccountName,LastBadPasswordAttempt,badPwdCount, LockedOut, DisplayName | Where{$_.LockedOut-eq $true} | Select-Object SamAccountName,LastBadPasswordAttempt,badPwdCount,DisplayName";
        }

        private void button76_Click(object sender, EventArgs e)
        {
            textBox76.Text = "Get-ADUser -Filter *  -Properties SamAccountName,LastBadPasswordAttempt,badPwdCount, LockedOut, DisplayName | Where{$_.LockedOut-eq $true} | Select-Object SamAccountName,LastBadPasswordAttempt,badPwdCount,DisplayName | Export-CSV \"H:\\AccountExpiryReport.csv\"";
        }



        //ADusers Report
        private void button53_Click(object sender, EventArgs e)
        {
            textBox78.Text = "Get-ADUser -Filter * -Properties * | Select -Property Name,GivenName,SamAccountName,Surname,Mail,Department,Company,enabled | export-csv H:\\SWSADUsers.csv";
        }

        
        private void button52_Click(object sender, EventArgs e)
        {
            textBox78.Text = "Get-ADUser -Filter * -SearchBase " + comboBox11.Text +" -Properties * | Select -Property Name,GivenName,SamAccountName,Surname,Mail,Department,Company,enabled | export-csv H:\\SWSADUsers.csv";
        }

        private void button51_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox78.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox78.Text);
            }
        }

        private void button50_Click(object sender, EventArgs e)
        {
            textBox78.Text = "";
        }
        // Membership Comparision
        private void button56_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox80.Text))
            {
                MessageBox.Show("Please Enter Logon Name1", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox81.Text))
            {
                MessageBox.Show("Please Enter Logon Name2", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (textBox80.Text == textBox81.Text)
            {
                MessageBox.Show("Please Enter Different Logon Names to Compare Membership", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox79.Text = "Compare-Object -ReferenceObject (Get-ADPrincipalGroupMembership -Identity " + textBox80.Text + ") -DifferenceObject (Get-ADPrincipalGroupMembership -Identity " + textBox81.Text + ") -Property Name | Sort Name";
            }
        }

        private void button55_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox79.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox79.Text);
            }
        }

        private void button54_Click(object sender, EventArgs e)
        {
            textBox79.Text = "";
            textBox80.Text = "";
            textBox81.Text = "";
        }

        private void button59_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox80.Text))
            {
                MessageBox.Show("Please Enter Logon Name1", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (string.IsNullOrEmpty(textBox81.Text))
            {
                MessageBox.Show("Please Enter Logon Name2", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if (textBox80.Text == textBox81.Text)
            {
                MessageBox.Show("Please Enter Different Logon Names to compare Membership", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox82.Text = "Compare-Object -ReferenceObject (Get-ADPrincipalGroupMembership -Identity " + textBox80.Text + ") -DifferenceObject (Get-ADPrincipalGroupMembership -Identity " + textBox81.Text + ") -Property Name | Sort Name | export-csv H:\\MembershipComparison.csv";
            }
        }
        
        private void button58_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox82.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox82.Text);
            }
        }

        private void button57_Click(object sender, EventArgs e)
        {
            textBox82.Text = "";
            textBox80.Text = "";
            textBox81.Text = "";
        }
        //AD groups export to excel
        private void button62_Click(object sender, EventArgs e)
        {
            textBox83.Text = "get-adgroup -filter * -properties Member | select Name,DistinguishedName,GroupCategory,GroupScope,Description | export-csv H:\\SWSADgroups.csv";
        }

        private void button61_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox83.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox83.Text);
            }
        }

        private void button60_Click(object sender, EventArgs e)
        {
            textBox83.Text = "";
        }
// Perticular AD Group users list
        private void button71_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox85.Text))
            {
                MessageBox.Show("Please AD Group Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox88.Text = "Get-ADGroupMember -identity " + textBox85.Text + " | select name,samaccountname | Export-csv H:\\Groupmembers.csv";
            }

        }

        private void button70_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox88.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            { 
                Clipboard.SetText(textBox88.Text); 
            }
        }

        private void button69_Click(object sender, EventArgs e)
        {
            textBox85.Text = "";
            textBox88.Text = "";
        }

        
        
        //Mailbox permission list
        private void button68_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox86.Text))
            {
                MessageBox.Show("Please Enter MailBox Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox87.Text = "Get-MailboxPermission -Identity " + textBox86.Text;
            }
        }

        private void button67_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox87.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox87.Text);
            }
        }

        private void button66_Click(object sender, EventArgs e)
        {
            textBox87.Text = "";
            textBox86.Text = "";
        }

        private void button65_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox86.Text))
            {
                MessageBox.Show("Please Enter MailBox Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox84.Text = "Get-MailboxPermission -Identity " + textBox86.Text + " | export-csv H:\\Mailboxpermission.csv";
            }
        }

        private void button64_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox84.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox84.Text);
            }
        }

        private void button63_Click(object sender, EventArgs e)
        {
            textBox84.Text = "";
            textBox86.Text = "";

        }
        //shared area owners
        private void button74_Click(object sender, EventArgs e)
        {
            textBox89.Text = "Get-ChildItem -Recurse | where {$_.psiscontainer} | Get-Acl | Export-Csv H:\\sharedareaowners.csv";
        }

        private void button73_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox89.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                Clipboard.SetText(textBox89.Text);
            }
        }

        private void button72_Click(object sender, EventArgs e)
        {
            textBox89.Text = "";
        }

        private void button79_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox92.Text))
            {
                MessageBox.Show("Please Enter Network Path To Find Owner", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                textBox91.Text = "Get-Acl \"" + textBox92.Text + "\"";
            }
            
        }

        private void button78_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox91.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {

                Clipboard.SetText(textBox91.Text);
            }
        }

        private void button77_Click(object sender, EventArgs e)
        {
            textBox91.Text = "";
            textBox92.Text = "";
        }
        //help button
        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HelpForm hfrm = new HelpForm();
            hfrm.Show();
        }
        //Application.Exit();
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Application.Exit();
            DialogResult dialog = MessageBox.Show("Do you really want to close PS Command Generator?", "Exit", MessageBoxButtons.YesNo);
            if (dialog == DialogResult.Yes)
            {
                Application.Exit();
            }
            else if (dialog == DialogResult.No)
            {
                e.Cancel = true;
            }
        }
        //Application.Exit();
        private void Exit_Click(object sender, EventArgs e)
        {
            DialogResult dialoga = MessageBox.Show("Do you really want to close the PS Command Generator?", "Exit", MessageBoxButtons.YesNo);
            if (dialoga == DialogResult.Yes)
            {
                Application.Exit();
            }
            else if (dialoga == DialogResult.No)
            {
                
            }
        }
        //AD Accounts Expiry Reports for 90 days
        private void button80_Click(object sender, EventArgs e)
        {
            textBox76.Text = "$90Days = (get-date).adddays(-90); Get-ADUser -properties * -filter {(lastlogondate -notlike \"*\" -OR lastlogondate -le $90days) -AND (passwordlastset -le $90days) -AND (enabled -eq $True) -and (PasswordNeverExpires -eq $false) -and (whencreated -le $90days)} | select-object name, SAMaccountname, passwordExpired, PasswordNeverExpires, logoncount, whenCreated, lastlogondate, PasswordLastSet | export-csv H:\\90days.CSV"; 
        }

        //AD Accounts Expiry Reports for 90 days from particular OU
        private void button83_Click(object sender, EventArgs e)
        {
            textBox90.Text = "Search-ADAccount -UsersOnly -SearchBase " + comboBox6.Text + " -AccountInactive -TimeSpan 90.00:00:00 | ?{$_.enabled -eq $true}  | Get-ADUser -Properties Name, sAMAccountName, Mail, Department, Company, enabled | export-csv H:\\unusedaccounts.csv -NoTypeInformation";
        }

        private void button82_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox90.Text))
            {
                MessageBox.Show("The Text Field is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {

                Clipboard.SetText(textBox90.Text);
            }
        }

        private void button81_Click(object sender, EventArgs e)
        {
            textBox90.Text = "";
        }

       
              
        
        
        
    }
}
//-AccessRights ReadProperty, WriteProperty -Properties 'Personal Information'";
//import-csv csvfile.csv | % {Add-ADGroupMember $_.groupname –Member username }